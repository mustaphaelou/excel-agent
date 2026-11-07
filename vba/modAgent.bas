Option Explicit
'==============================================================================
' modAgent - panel host: ribbon entry points, the modeless form, the layout
' timer, and the self-tests. Conversation state lives in Turn.bas.
'
' The form is created from the designer at runtime (UserForms.Add) so this
' module never holds a compile-time reference to it - that keeps the add-in
' compiling even if the form component is rebuilt.
'==============================================================================

Public Const APP_NAME As String = "Excel Agent"


'--- form instance ------------------------------------------------------------
Public gForm As Object

'--- layout poll state --------------------------------------------------------
' Must live up here with the other module-level variables. VBA rejects a
' declaration that appears after a procedure with "Only comments may appear
' after End Sub, End Function, or End Property" - and because VBA compiles on
' demand, that error stays invisible until something calls into the module.
Private gLayoutRunning As Boolean


'==============================================================================
' Trace - append-only breadcrumbs for the self-test. Written to the Windows
' temp folder because WSL paths over 9p are unreliable from inside VBA.
'==============================================================================
Public Sub Trace(ByVal s As String)
    Dim fn As Integer, p As String
    On Error Resume Next
    p = Environ$("TEMP") & "\excelagent_trace.log"
    fn = FreeFile
    Open p For Append As #fn
    Print #fn, Format$(Now, "hh:mm:ss") & "  " & s
    Close #fn
End Sub


'==============================================================================
' Layout loop - MSForms does not reliably fire UserForm_Resize while the user
' drags a border, so while the panel is visible we re-check its size once a
' second. ApplyLayout is a no-op unless the size actually changed.
'==============================================================================
Public Sub StartLayoutLoop()
    If gLayoutRunning Then Exit Sub
    gLayoutRunning = True
    ScheduleLayout
End Sub

Public Sub StopLayoutLoop()
    gLayoutRunning = False          ' the next tick sees the flag and stops
End Sub

Private Sub ScheduleLayout()
    On Error Resume Next
    Application.OnTime Now + TimeSerial(0, 0, 1), "FormLayoutTick"
End Sub

Public Sub FormLayoutTick()
    On Error Resume Next
    If Not gLayoutRunning Then Exit Sub
    If gForm Is Nothing Then
        gLayoutRunning = False
        Exit Sub
    End If
    gForm.CheckSize
    If gForm.Visible Then
        ScheduleLayout
    Else
        gLayoutRunning = False
    End If
End Sub


'==============================================================================
' Ribbon entry points
'==============================================================================
Public Sub AgentShow(Optional ByVal control As Object)
    ShowAgentForm ""
End Sub

Public Sub AgentAskSelection(Optional ByVal control As Object)
    ShowAgentForm "Explain what this selected data shows, then suggest one improvement."
End Sub

Public Sub AgentSummarizeSheet(Optional ByVal control As Object)
    ShowAgentForm "Summarise the active sheet in three bullets."
End Sub

Public Sub AgentReset(Optional ByVal control As Object)
    If Turn.Busy() Then
        MsgBox "Wait for the current request to finish.", vbInformation, APP_NAME
        Exit Sub
    End If
    If Not gForm Is Nothing Then
        If gForm.Busy Then
            MsgBox "Wait for the current request to finish.", vbInformation, APP_NAME
            Exit Sub
        End If
    End If
    Turn.Clear
    If Not gForm Is Nothing Then
        Unload gForm
        Set gForm = Nothing
    End If
    MsgBox "Conversation cleared.", vbInformation, APP_NAME
End Sub


Private Sub ShowAgentForm(ByVal prefill As String)
    On Error GoTo Fail
    If gForm Is Nothing Then
        Set gForm = VBA.UserForms.Add("frmAgent")
    End If
    If Not gForm.Visible Then gForm.Show vbModeless
    gForm.Prefill prefill
    Exit Sub
Fail:
    MsgBox "Could not open the agent panel: " & Err.Description, vbExclamation, APP_NAME
End Sub

' Keep the panel usable if the user closes it with the X.
Public Sub FormClosed()
    Set gForm = Nothing
End Sub


'==============================================================================
' FormSelfTest - verifies the runtime-built UserForm actually instantiates,
' builds its controls, shows modeless, and tears down. The main self-test does
' not touch the UI; this covers the code path a real click takes.
'==============================================================================
Public Function FormSelfTest(ByVal outPath As String) As String
    Dim sb As String, fn As Integer

    On Error GoTo Failed
    sb = "form self-test" & vbCrLf
    Trace "formtest: creating"
    Set gForm = VBA.UserForms.Add("frmAgent")
    sb = sb & "instantiated: yes" & vbCrLf
    sb = sb & "controls: " & gForm.Controls.Count & vbCrLf
    sb = sb & "model items: " & gForm.ModelItemCount & vbCrLf
    sb = sb & "model selected: " & gForm.SelectedModelId & vbCrLf

    Trace "formtest: showing"
    gForm.Show vbModeless
    DoEvents
    sb = sb & "visible: " & gForm.Visible & vbCrLf
    sb = sb & "model combo: " & gForm.ModelComboRect & " index=" & gForm.ModelListIndex & vbCrLf

    ' resizing: exercise the layout engine the way a drag would
    Dim ow As Single, oh As Single
    ow = Val(GetSetting("ExcelAgent", "Panel", "Width", "442"))
    oh = Val(GetSetting("ExcelAgent", "Panel", "Height", "566"))
    If ow < 380 Then ow = 442
    If oh < 340 Then oh = 566

    Trace "formtest: resizing"
    gForm.ResizeTo 780, 820
    DoEvents
    sb = sb & "resizable style: " & gForm.ResizableApplied & vbCrLf
    sb = sb & "inside after resize: " & Int(gForm.InsideWidth) & "x" & Int(gForm.InsideHeight) & vbCrLf
    sb = sb & "log box after resize: " & Int(gForm.LogWidth) & "x" & Int(gForm.LogHeight) & vbCrLf
    sb = sb & "model combo after resize: " & gForm.ModelComboRect & vbCrLf
    sb = sb & "layout expanded: " & (gForm.LogWidth > 700 And gForm.LogHeight > 300) & vbCrLf

    gForm.ResizeTo ow, oh
    DoEvents
    sb = sb & "log box restored: " & Int(gForm.LogWidth) & "x" & Int(gForm.LogHeight) & vbCrLf

    Trace "formtest: hiding"
    gForm.Hide
    DoEvents
    Unload gForm
    Set gForm = Nothing
    Trace "formtest: done"

    sb = sb & "RESULT: PASS"
    GoTo WriteIt

Failed:
    sb = sb & "EXCEPTION: " & Err.Description & vbCrLf & "RESULT: FAIL"

WriteIt:
    On Error Resume Next
    fn = FreeFile
    Open outPath For Output As #fn
    Print #fn, sb
    Close #fn
    FormSelfTest = sb
End Function


'==============================================================================
' SelfTest - driven by build.ps1 so the whole stack gets exercised before
' anyone touches a keyboard.
'==============================================================================
Public Function SelfTest(ByVal outPath As String) As String
    Dim sb As String, status As String, ok As Boolean
    Dim result As TurnResult
    Dim result2 As TurnResult
    Dim wb As Workbook, wb2 As Workbook, ws As Worksheet, res As String
    Dim label As Variant
    Dim fnum As Integer

    On Error GoTo Failed
    Trace "selftest: start"
    sb = "Excel " & Application.Version & " / build " & Application.Build & vbCrLf
    sb = sb & "Bridge reachable: " & BridgeAlive() & vbCrLf

    Trace "selftest: clearing turn"
    Turn.Clear

    Trace "selftest: ensure bridge"
    ok = EnsureBridge(status)
    sb = sb & "Bridge: " & ok & " (" & status & ")" & vbCrLf

    Trace "selftest: model catalog"
    ok = LoadModels(True)
    sb = sb & "Models: " & ModelCount() & " (live: " & ModelsFromBridge() & ")" & vbCrLf
    sb = sb & "Selected model: " & ActiveModel() & vbCrLf

    ' scratch workbook so the test never touches user data
    Trace "selftest: adding scratch workbook"
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Sales"
    ws.Range("A1:C1").Value2 = Array("Region", "Units", "Revenue")
    ws.Range("A2").Value2 = "North": ws.Range("B2").Value2 = 120: ws.Range("C2").Value2 = 4500
    ws.Range("A3").Value2 = "South": ws.Range("B3").Value2 = 80:  ws.Range("C3").Value2 = 3100
    wb.Activate
    Trace "selftest: seeding data"
    ws.Range("A1:C3").Select

    Trace "selftest: sending prompt"
    Set result = Turn.Submit("Add a Total column in D that is Units times Revenue, and bold the header row.", _
                             False, ActiveModel())
    Trace "selftest: reply received ok=" & result.Ok & " edits=" & result.Edits.Count
    sb = sb & "Send ok: " & result.Ok & " err=" & result.ErrorText & vbCrLf
    sb = sb & "Status: " & result.Status & vbCrLf
    sb = sb & "Reply: " & Replace(result.Reply, vbLf, " ") & vbCrLf
    sb = sb & "Edits: " & result.Edits.Count & vbCrLf
    For Each label In result.Edits
        sb = sb & "  " & label & vbCrLf
    Next label
    sb = sb & "Pending after submit: " & Turn.HasPending() & vbCrLf
    sb = sb & "Pending count: " & Turn.PendingCount() & vbCrLf
    sb = sb & "Pending summary: " & Replace(Turn.PendingSummary(), vbCrLf, " | ") & vbCrLf

    If Not result.Ok Then
        sb = sb & "RESULT: FAIL" & vbCrLf
        GoTo Finish
    End If

    Trace "selftest: blocked second send"
    Set result2 = Turn.Submit("This second request must wait for the pending edits.", False, ActiveModel())
    sb = sb & "Second send blocked: " & (Not result2.Ok) & " (" & result2.ErrorText & ")" & vbCrLf
    If result2.Ok Then
        sb = sb & "RESULT: FAIL (pending edits did not block the new turn)" & vbCrLf
        GoTo Finish
    End If

    Trace "selftest: wrong-workbook refusal"
    Set wb2 = Workbooks.Add
    wb2.Activate
    res = Turn.Apply()
    sb = sb & "Wrong-workbook apply: " & res & vbCrLf
    sb = sb & "Still pending after refusal: " & Turn.HasPending() & vbCrLf
    If InStr(1, res, "Activate", vbTextCompare) = 0 Then
        sb = sb & "RESULT: FAIL (wrong workbook was not refused)" & vbCrLf
        GoTo Finish
    End If
    If Not Turn.HasPending() Then
        sb = sb & "RESULT: FAIL (refusal lost the pending edits)" & vbCrLf
        GoTo Finish
    End If

    wb2.Close SaveChanges:=False
    Set wb2 = Nothing
    wb.Activate

    Trace "selftest: applying"
    res = Turn.Apply()
    Trace "selftest: applied"
    sb = sb & "Apply: " & res & vbCrLf
    sb = sb & "Pending after apply: " & Turn.HasPending() & vbCrLf
    sb = sb & "D2 value : " & ws.Range("D2").Formula & " => " & ws.Range("D2").Value2 & vbCrLf
    sb = sb & "D4 empty : " & (Len(CStr(ws.Range("D4").Value2)) = 0) & vbCrLf
    sb = sb & "A1 bold  : " & ws.Range("A1").Font.Bold & vbCrLf
    sb = sb & "History  : " & Turn.HistoryCount() & vbCrLf

    If Turn.HasPending() Then
        sb = sb & "RESULT: FAIL (edits still pending after apply)" & vbCrLf
        GoTo Finish
    End If
    If Turn.HistoryCount() <> 2 Then
        sb = sb & "RESULT: FAIL (history count " & Turn.HistoryCount() & ")" & vbCrLf
        GoTo Finish
    End If

    sb = sb & "RESULT: PASS"
    GoTo Finish

Failed:
    Trace "selftest: EXCEPTION " & Err.Description
    sb = sb & "EXCEPTION: " & Err.Description & vbCrLf & "RESULT: FAIL"

Finish:
    On Error Resume Next
    If Not wb2 Is Nothing Then wb2.Close SaveChanges:=False
    Set wb2 = Nothing
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wb = Nothing

    fnum = FreeFile
    Open outPath For Output As #fnum
    Print #fnum, sb
    Close #fnum
    SelfTest = sb
End Function
