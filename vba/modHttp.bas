Option Explicit
'==============================================================================
' modHttp - UTF-8 safe HTTP against the local bridge.
'
' WinHTTP is used (not MSXML) because it needs no COM registration beyond what
' ships with Windows. Bodies are converted through ADODB.Stream so Arabic,
' French accents, and emoji survive the trip in both directions.
'
' The async variant pumps DoEvents while waiting, so a slow model call does not
' leave Excel looking hung.
'==============================================================================

Public Const BRIDGE_HOST As String = "http://127.0.0.1:8765"
' The bridge script path lives here, next to the process that starts it.
Public Const BRIDGE_SCRIPT As String = "/home/eloaurdi/hay2010agent/excel-agent/bridge.py"

#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

'--- encoding helpers ---------------------------------------------------------
Private Function Utf8Bytes(ByVal s As String) As Variant
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2          ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText s
    stm.Position = 0
    stm.Type = 1          ' adTypeBinary
    Utf8Bytes = stm.Read
    stm.Close
    Set stm = Nothing
End Function

Private Function BytesToText(ByVal b As Variant) As String
    Dim stm As Object
    If IsEmpty(b) Or IsNull(b) Then Exit Function
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write b
    stm.Position = 0
    stm.Type = 2
    stm.Charset = "utf-8"
    BytesToText = stm.ReadText
    stm.Close
    Set stm = Nothing
End Function

'--- health -------------------------------------------------------------------
Public Function BridgeAlive() As Boolean
    Dim http As Object
    On Error GoTo Dead
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", BRIDGE_HOST & "/health", False
    http.SetTimeouts 1500, 1500, 1500, 2500
    http.Send
    BridgeAlive = (http.Status = 200)
    Exit Function
Dead:
    Trace "  alive: failed - " & Err.Description
    BridgeAlive = False
End Function

'--- GET (synchronous, short timeout) -----------------------------------------
' Used for the model catalog, which must answer quickly or not at all.
Public Function GetText(ByVal path As String, ByRef errMsg As String) As String
    Dim http As Object
    errMsg = ""
    On Error GoTo Fail
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", BRIDGE_HOST & path, False
    http.SetTimeouts 1500, 1500, 3000, 8000
    http.SetRequestHeader "Accept", "text/plain"
    http.Send
    If http.Status <> 200 Then
        errMsg = "HTTP " & http.Status & " " & http.StatusText
        Exit Function
    End If
    GetText = BytesToText(http.ResponseBody)
    Exit Function
Fail:
    errMsg = "Cannot reach the agent bridge (" & Err.Description & ")."
End Function

'--- POST (async, pumps the message loop) -------------------------------------
Public Function Post(body As String, ByRef errMsg As String, _
                     Optional ByVal timeoutSec As Long = 180) As String
    Dim http As Object
    Dim t0 As Single, waited As Single
    errMsg = ""

    On Error GoTo Fail
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", BRIDGE_HOST & "/chat", True
    http.SetTimeouts 3000, 3000, 20000, timeoutSec * 1000
    http.SetRequestHeader "Content-Type", "text/plain; charset=utf-8"
    http.SetRequestHeader "Accept", "text/plain"
    Trace "http: sending "
    http.Send Utf8Bytes(body)

    t0 = Timer
    Do While Not http.WaitForResponse(1)
        DoEvents
        waited = Timer - t0
        If waited < 0 Then waited = waited + 86400   ' midnight rollover
        If waited > timeoutSec Then
            errMsg = "Timed out after " & timeoutSec & "s waiting for the bridge."
            On Error Resume Next
            http.Abort
            Exit Function
        End If
    Loop

    If http.Status <> 200 Then
        errMsg = "HTTP " & http.Status & " " & http.StatusText
        Exit Function
    End If
    Trace "http: got " & http.Status
    Post = BytesToText(http.ResponseBody)
    Exit Function

Fail:
    errMsg = "Cannot reach the agent bridge at " & BRIDGE_HOST & " (" & Err.Description & ")."
End Function

'--- bring the bridge up if it is not running ---------------------------------
Public Function EnsureBridge(ByRef statusMsg As String) As Boolean
    Dim sh As String, i As Long
    If BridgeAlive() Then
        statusMsg = "Connected."
        EnsureBridge = True
        Exit Function
    End If

    statusMsg = "Starting agent service..."
    On Error Resume Next
    sh = "wsl.exe -d Ubuntu -- python3 " & BRIDGE_SCRIPT & " --port 8765"
    Shell sh, vbHide
    On Error GoTo 0

    For i = 1 To 40           ' up to ~10s
        DoEvents
        Sleep 250
        If BridgeAlive() Then
            EnsureBridge = True
            statusMsg = "Agent service ready."
            Exit Function
        End If
    Next
    statusMsg = "Agent service did not start. Run bridge.py manually."
    EnsureBridge = False
End Function
