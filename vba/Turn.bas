Option Explicit
'==============================================================================
' Turn - one conversational turn: history, pending edits, bridge call, apply.
'
' The panel is a view over this module. It owns:
'   * the conversation history (MAX_HIST messages; the bridge sees the last 12)
'   * the pending edits, their display labels, and the workbook they belong to
'   * the bridge call (ensure, post, parse)
'   * apply / discard / clear and the busy flag
'
' It never touches the form, the ribbon, or the model picker. Errors come back
' in TurnResult; nothing here raises a modal.
'
' VBA has no exceptions, so every public entry point catches its own errors
' and returns a message. A pending edit stays pending until it is applied or
' discarded: a failed apply does not lose the proposal.
'==============================================================================

Private Const MAX_HIST As Long = 20

Private mHistRole() As String
Private mHistText() As String
Private mHistCount As Long

Private mPending As Collection          ' each item is a String() ACTION record
Private mPendingLabels As Collection    ' parallel display labels
Private mPendingWorkbook As String
Private mBusy As Boolean


'==============================================================================
' Public surface
'==============================================================================
Public Function Submit(ByVal userText As String, ByVal readonlyMode As Boolean, _
                       ByVal modelId As String) As TurnResult
    Dim result As TurnResult
    Dim wb As Workbook
    Dim body As String, resp As String, errMsg As String
    Dim status As String
    Dim recs As Collection, p() As String, i As Long, kind As String
    Dim reply As String, label As String
    Dim pending As Collection, labels As Collection

    Set result = New TurnResult
    On Error GoTo Failed

    If mBusy Then
        result.ErrorText = "A turn is already running."
        Set Submit = result
        Exit Function
    End If
    If HasPending() Then
        result.ErrorText = "Apply or discard the proposed edits first."
        Set Submit = result
        Exit Function
    End If

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        result.ErrorText = "Open a workbook first."
        Set Submit = result
        Exit Function
    End If
    If StrComp(wb.Name, ThisWorkbook.Name, vbTextCompare) = 0 Then
        result.ErrorText = "Activate your own workbook, then ask again."
        Set Submit = result
        Exit Function
    End If

    mBusy = True

    If Not EnsureBridge(status) Then
        result.ErrorText = status
        result.Status = status
        mBusy = False
        Set Submit = result
        Exit Function
    End If
    result.Status = status
    If status = "Agent service ready." Then
        result.Notes.Add "Agent service started."
    End If

    body = BuildRequest(userText, readonlyMode, modelId, mHistRole, mHistText, mHistCount)
    resp = Post(body, errMsg)
    If Len(errMsg) > 0 Then
        result.ErrorText = errMsg
        mBusy = False
        Set Submit = result
        Exit Function
    End If

    Set recs = ParseResponse(resp)
    Set pending = New Collection
    Set labels = New Collection
    For i = 1 To recs.Count
        p = recs(i)
        kind = KindOf(p)
        Select Case kind
            Case "REPLY"
                reply = reply & FieldOr(p, 1, "")
            Case "NOTE"
                result.Notes.Add FieldOr(p, 1, "")
            Case "ACTION"
                label = "[" & FieldOr(p, 1, "") & "] " & FieldOr(p, 2, "")
                labels.Add label
                pending.Add p
        End Select
    Next i

    If Len(reply) = 0 Then reply = "(the agent returned no text)"
    result.Reply = reply
    result.Ok = True
    Set result.Edits = labels

    If pending.Count > 0 Then
        Set mPending = pending
        Set mPendingLabels = labels
        mPendingWorkbook = wb.Name
    Else
        Set mPending = Nothing
        Set mPendingLabels = Nothing
        mPendingWorkbook = ""
    End If

    AddHistory "user", userText
    AddHistory "assistant", reply

    mBusy = False
    Set Submit = result
    Exit Function

Failed:
    If result Is Nothing Then Set result = New TurnResult
    result.Ok = False
    result.ErrorText = "Turn failed: " & Err.Description
    mBusy = False
    Set Submit = result
End Function


Public Function Apply() As String
    Dim wb As Workbook
    Dim i As Long, applied As Long, failed As Long
    Dim msg As String, res As String
    Dim p() As String
    Dim kept As Collection, keptLabels As Collection

    If mBusy Then
        Apply = "A turn is already running."
        Exit Function
    End If
    If Not HasPending() Then
        Apply = "Nothing to apply."
        Exit Function
    End If

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Apply = "No workbook is active."
        Exit Function
    End If
    If StrComp(wb.Name, mPendingWorkbook, vbTextCompare) <> 0 Then
        Apply = "Activate '" & mPendingWorkbook & "' and try again."
        Exit Function
    End If

    mBusy = True
    On Error GoTo Failed

    Set kept = New Collection
    Set keptLabels = New Collection

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For i = 1 To mPending.Count
        p = mPending(i)
        res = ApplyOne(p, wb)
        If Len(res) = 0 Then
            applied = applied + 1
        Else
            failed = failed + 1
            kept.Add p
            keptLabels.Add mPendingLabels(i)
            msg = msg & vbCrLf & "  - " & mPendingLabels(i) & " (" & res & ")"
        End If
    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    If kept.Count > 0 Then
        Set mPending = kept
        Set mPendingLabels = keptLabels
    Else
        Set mPending = Nothing
        Set mPendingLabels = Nothing
        mPendingWorkbook = ""
    End If

    mBusy = False
    Apply = apply_report(applied, failed, msg)
    Exit Function

Failed:
    On Error Resume Next
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    mBusy = False
    Apply = "Apply failed: " & Err.Description
End Function


Private Function apply_report(ByVal applied As Long, ByVal failed As Long, _
                              ByVal msg As String) As String
    If failed = 0 Then
        apply_report = "Applied " & applied & " edit" & IIf(applied = 1, "", "s") & "."
    ElseIf applied = 0 Then
        apply_report = "No edits applied; " & failed & " still pending:" & msg
    Else
        apply_report = "Applied " & applied & "; " & failed & " still pending:" & msg
    End If
End Function


Public Sub Discard()
    If mBusy Then Exit Sub
    Set mPending = Nothing
    Set mPendingLabels = Nothing
    mPendingWorkbook = ""
End Sub


Public Sub Clear()
    If mBusy Then Exit Sub
    mHistCount = 0
    ReDim mHistRole(1 To MAX_HIST)
    ReDim mHistText(1 To MAX_HIST)
    Set mPending = Nothing
    Set mPendingLabels = Nothing
    mPendingWorkbook = ""
End Sub


Public Function HasPending() As Boolean
    If mPending Is Nothing Then Exit Function
    HasPending = (mPending.Count > 0)
End Function


Public Function PendingCount() As Long
    If mPending Is Nothing Then Exit Function
    PendingCount = mPending.Count
End Function


Public Function PendingSummary() As String
    Dim i As Long, s As String
    If Not HasPending() Then
        PendingSummary = "(none)"
        Exit Function
    End If
    For i = 1 To mPendingLabels.Count
        If Len(s) > 0 Then s = s & vbCrLf
        s = s & "  " & i & ". " & mPendingLabels(i)
    Next i
    PendingSummary = s
End Function


Public Function HistoryCount() As Long
    HistoryCount = mHistCount
End Function


Public Function Busy() As Boolean
    Busy = mBusy
End Function


'==============================================================================
' History - the same 20-slot FIFO the add-in had before, now owned here.
'==============================================================================
Private Sub AddHistory(ByVal role As String, ByVal text As String)
    Dim i As Long
    If mHistCount = 0 Then
        ReDim mHistRole(1 To MAX_HIST)
        ReDim mHistText(1 To MAX_HIST)
    End If

    If mHistCount >= MAX_HIST Then
        For i = 2 To MAX_HIST
            mHistRole(i - 1) = mHistRole(i)
            mHistText(i - 1) = mHistText(i)
        Next i
        mHistCount = MAX_HIST - 1
    End If

    mHistCount = mHistCount + 1
    mHistRole(mHistCount) = role
    mHistText(mHistCount) = text
End Sub
