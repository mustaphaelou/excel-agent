Option Explicit
'==============================================================================
' modProbe - breadcrumb probes used by build/loop.sh to bisect a stall.
'
' Every step appends a line to the file passed in as `p`, BEFORE attempting the
' call, so a hang still leaves evidence of the last statement reached.
'
' NAMES: never use <letters><digits> (e.g. S1, P2). Application.Run resolves
' such a string as a cell reference and fails with 0x800A03EC before the macro
' runs at all. Letters only, or an underscore, is safe.
'==============================================================================

Public Sub PbLog(ByVal p As String, ByVal s As String)
    Dim fn As Integer
    On Error Resume Next
    fn = FreeFile
    Open p For Append As #fn
    Print #fn, s
    Close #fn
End Sub

'--- baseline: nothing but our own log writer --------------------------------
Public Sub PbNull(ByVal p As String)
    PbLog p, "null: start"
    PbLog p, "null: done"
End Sub

'--- host object model only --------------------------------------------------
Public Sub PbApp(ByVal p As String)
    PbLog p, "app: start"
    PbLog p, "app: version=" & Application.Version
    PbLog p, "app: done"
End Sub

'--- cross-module: modAgent Trace --------------------------------------------
Public Sub PbTrace(ByVal p As String)
    PbLog p, "trace: start"
    Trace "probe: via modAgent"
    PbLog p, "trace: after Trace"
    PbLog p, "trace: done"
End Sub

'--- cross-module: modAgent ResetHistory -------------------------------------
Public Sub PbReset(ByVal p As String)
    PbLog p, "reset: start"
    ResetHistory
    PbLog p, "reset: after ResetHistory"
    PbLog p, "reset: done"
End Sub

'--- cross-module: modWindow (pure Win32) ------------------------------------
Public Sub PbWindow(ByVal p As String)
    Dim ok As Boolean
    PbLog p, "window: start"
    ok = MakeResizable("no such window")
    PbLog p, "window: ok=" & ok
    PbLog p, "window: done"
End Sub

'--- cross-module: modHttp BridgeAlive ---------------------------------------
Public Sub PbAlive(ByVal p As String)
    Dim alive As Boolean
    PbLog p, "alive: start"
    alive = BridgeAlive()
    PbLog p, "alive: alive=" & alive
    PbLog p, "alive: done"
End Sub

'--- cross-module: modHttp EnsureBridge --------------------------------------
Public Sub PbEnsure(ByVal p As String)
    Dim st As String, r As Boolean
    PbLog p, "ensure: start"
    r = EnsureBridge(st)
    PbLog p, "ensure: r=" & r & " st=" & st
    PbLog p, "ensure: done"
End Sub

'--- THE SUSPECT: EnsureBridgeOrWarn -----------------------------------------
Public Sub PbGuard(ByVal p As String)
    Dim st As String, r As Boolean
    PbLog p, "guard: start"
    r = EnsureBridgeOrWarn(st)
    PbLog p, "guard: r=" & r & " st=" & st
    PbLog p, "guard: done"
End Sub

'--- mirrors SelfTest's order: alive first, then guard -----------------------
Public Sub PbOrder(ByVal p As String)
    Dim st As String, r As Boolean
    PbLog p, "order: start"
    r = BridgeAlive()
    PbLog p, "order: alive=" & r
    r = EnsureBridgeOrWarn(st)
    PbLog p, "order: guard=" & r
    PbLog p, "order: done"
End Sub

'--- same body as EnsureBridgeOrWarn but with no On Error handler ------------
Public Function PbGuardNoHandler(ByRef st As String) As Boolean
    If BridgeAlive() Then
        PbGuardNoHandler = True
        st = "Connected."
        Exit Function
    End If
    PbGuardNoHandler = EnsureBridge(st)
End Function

Public Sub PbNoHandler(ByVal p As String)
    Dim st As String, r As Boolean
    PbLog p, "nohandler: start"
    r = PbGuardNoHandler(st)
    PbLog p, "nohandler: r=" & r & " st=" & st
    PbLog p, "nohandler: done"
End Sub

'--- form layout: exact control geometry at two sizes -----------------------
Public Sub PbFormLayout(ByVal p As String)
    Dim f As Object, c As Object
    Set f = VBA.UserForms.Add("frmAgent")
    f.Show vbModeless
    DoEvents
    PbLog p, "form inside=" & Int(f.InsideWidth) & "x" & Int(f.InsideHeight)
    DumpControls p, f
    f.ResizeTo 880, 700
    DoEvents
    PbLog p, "-- after ResizeTo 880x700 --"
    PbLog p, "form inside=" & Int(f.InsideWidth) & "x" & Int(f.InsideHeight)
    DumpControls p, f
    f.Hide
    Unload f
End Sub

Private Sub DumpControls(ByVal p As String, ByVal f As Object)
    Dim c As Object
    For Each c In f.Controls
        PbLog p, "  " & Left$(c.Name & Space$(14), 14) & _
                 " " & Int(c.Left) & "," & Int(c.Top) & _
                 " " & Int(c.Width) & "x" & Int(c.Height)
    Next
End Sub
