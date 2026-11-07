Option Explicit
'==============================================================================
' modModels - the model picker's data and the current selection.
'
' bridge.py owns the catalog (GET /models): it asks the gateway, caches the
' answer, and falls back to a static list. VBA keeps its own copy here so the
' panel opens instantly and works with the bridge down.
'
' The choice is per user, not per workbook - switching workbooks keeps it.
'==============================================================================

Public Const DEFAULT_MODEL As String = "beam-free"

' Shown until /models answers. Keep in sync with FALLBACK_CATALOG in bridge.py.
Private Const FALLBACK_IDS As String = _
    "beam-free|deepseek-v4.1-flash|glm-5.3-flash|qwen3.8-flash|gpt-6-luna|claude-haiku-5-5"
Private Const FALLBACK_LABELS As String = _
    "Beam Free|DeepSeek V4.1 Flash|GLM-5.3-Flash|Qwen3.8 Flash|GPT-6 Luna|Claude Haiku 5.5"

Private mIds() As String
Private mLabels() As String
Private mCount As Long
Private mFromBridge As Boolean
Private mSelected As String
Private mRefreshQueued As Boolean

'==============================================================================
' Catalog access (1-based, matching the combo's ListIndex + 1)
'==============================================================================
Public Function ModelCount() As Long
    ModelCount = mCount
End Function

Public Function ModelIdAt(ByVal idx As Long) As String
    If idx >= 1 And idx <= mCount Then ModelIdAt = mIds(idx)
End Function

Public Function ModelLabelAt(ByVal idx As Long) As String
    If idx >= 1 And idx <= mCount Then ModelLabelAt = mLabels(idx)
End Function

Public Function ModelIndexById(ByVal id As String) As Long
    Dim i As Long
    If Len(id) = 0 Then Exit Function
    For i = 1 To mCount
        If StrComp(mIds(i), id, vbTextCompare) = 0 Then
            ModelIndexById = i
            Exit Function
        End If
    Next i
End Function

' True when the current list came from the gateway, not from the static fallback.
Public Function ModelsFromBridge() As Boolean
    ModelsFromBridge = mFromBridge
End Function

'==============================================================================
' Selection
'==============================================================================
Public Sub EnsureModels()
    If mCount = 0 Then FillFallback

    If Len(mSelected) = 0 Then mSelected = SavedModel()
    If Len(mSelected) = 0 Then mSelected = DEFAULT_MODEL

    ' The saved choice may be missing from the small offline list, because the
    ' live catalog is much bigger. Keep it visible until /models answers;
    ' LoadModels decides whether it still exists upstream.
    If ModelIndexById(mSelected) = 0 Then AddModel mSelected, mSelected
End Sub

Public Function ActiveModel() As String
    EnsureModels
    ActiveModel = mSelected
End Function

Public Function ActiveModelLabel() As String
    Dim idx As Long
    EnsureModels
    idx = ModelIndexById(ActiveModel())
    If idx = 0 Then idx = 1
    ActiveModelLabel = ModelLabelAt(idx)
End Function

Public Sub SetActiveModel(ByVal id As String)
    EnsureModels
    If ModelIndexById(id) = 0 Then id = DEFAULT_MODEL
    mSelected = id
    On Error Resume Next
    SaveSetting "ExcelAgent", "Model", "Id", id
End Sub

Private Function SavedModel() As String
    On Error Resume Next
    SavedModel = Trim$(GetSetting("ExcelAgent", "Model", "Id", DEFAULT_MODEL))
    If Len(SavedModel) = 0 Then SavedModel = DEFAULT_MODEL
End Function

Private Sub AddModel(ByVal id As String, ByVal label As String)
    If Len(id) = 0 Then Exit Sub
    mCount = mCount + 1
    ReDim Preserve mIds(1 To mCount)
    ReDim Preserve mLabels(1 To mCount)
    mIds(mCount) = id
    mLabels(mCount) = label
End Sub

Private Sub FillFallback()
    Dim ids As Variant, labels As Variant, i As Long
    ' Split() is 0-based; every other array here is 1-based, matching the
    ' combo's ListIndex + 1. Copy across rather than assigning Split directly.
    ids = Split(FALLBACK_IDS, "|")
    labels = Split(FALLBACK_LABELS, "|")
    mCount = UBound(ids) + 1
    ReDim mIds(1 To mCount)
    ReDim mLabels(1 To mCount)
    For i = 0 To UBound(ids)
        mIds(i + 1) = CStr(ids(i))
        mLabels(i + 1) = CStr(labels(i))
    Next i
    mFromBridge = False
End Sub

'==============================================================================
' Catalog refresh
'==============================================================================
' Returns True when the bridge answered with a usable catalog (live or
' fallback); the list is then safe to display. Any failure leaves the current
' list untouched.
Public Function LoadModels(ByVal fromBridge As Boolean) As Boolean
    Dim resp As String, errMsg As String, recs As Collection
    Dim p() As String, i As Long, n As Long
    Dim ids() As String, labels() As String
    Dim live As Boolean, loaded As Boolean

    EnsureModels                    ' settle the selection before replacing the list
    If fromBridge Then
        resp = GetText("/models", errMsg)
        If Len(errMsg) = 0 And Len(resp) > 0 Then
            Set recs = ParseResponse(resp)
            If recs.Count > 0 Then
                ReDim ids(1 To recs.Count)
                ReDim labels(1 To recs.Count)
                For i = 1 To recs.Count
                    p = recs(i)
                    If KindOf(p) = "SOURCE" Then
                        live = (LCase$(FieldOr(p, 1, "")) = "live")
                    ElseIf KindOf(p) = "MODEL" Then
                        If Len(FieldOr(p, 1, "")) > 0 Then
                            n = n + 1
                            ids(n) = FieldOr(p, 1, "")
                            labels(n) = FieldOr(p, 2, ids(n))
                        End If
                    End If
                Next i
                If n > 0 Then
                    ReDim Preserve ids(1 To n)
                    ReDim Preserve labels(1 To n)
                    ' A fallback answer must not replace a list we already got
                    ' live: the combo still shows those entries, and swapping
                    ' the arrays under it would map the wrong model to a row.
                    If live Or Not mFromBridge Then
                        mIds = ids
                        mLabels = labels
                        mCount = n
                        mFromBridge = live
                    End If
                    loaded = True
                End If
            End If
        End If
    End If

    If mCount = 0 Then FillFallback

    ' Only a live catalog is authoritative. If it no longer offers the saved
    ' choice, fall back (default first, then whatever the list starts with) and
    ' persist that so the next open agrees with the combo. A fallback catalog
    ' must not delete the selection: the gateway may just be briefly down.
    If mFromBridge And ModelIndexById(mSelected) = 0 Then
        mSelected = SavedModel()
        If ModelIndexById(mSelected) = 0 Then
            If mCount > 0 Then mSelected = mIds(1) Else mSelected = DEFAULT_MODEL
            On Error Resume Next
            SaveSetting "ExcelAgent", "Model", "Id", mSelected
        End If
    End If

    EnsureModels                    ' keeps an offline saved choice visible
    LoadModels = loaded
End Function

' Fetches the live list one second after the panel opens, but only when the
' bridge is already running. Starting the bridge here would freeze the UI for
' up to ten seconds; the first Send does that (and reschedules a refresh).
Public Sub ScheduleModelRefresh()
    If mRefreshQueued Then Exit Sub
    On Error GoTo Failed
    mRefreshQueued = True
    Application.OnTime Now + TimeSerial(0, 0, 1), "RefreshModelsTick"
    Exit Sub
Failed:
    mRefreshQueued = False
End Sub

Public Sub RefreshModelsTick()
    mRefreshQueued = False
    On Error Resume Next
    If Not BridgeAlive() Then Exit Sub
    If Not LoadModels(True) Then Exit Sub
    If gForm Is Nothing Then Exit Sub
    gForm.RefreshModelCombo
End Sub
