Option Explicit
'==============================================================================
' modContext - turns the live workbook into the snapshot the model reads.
'
' Token budget matters: every sheet is capped at 80x26 cells, cells at 160
' characters, and the whole snapshot at ~40k characters. Anything bigger and
' the first response of the session gets slow and expensive.
'==============================================================================

Private Const MAX_ROWS As Long = 80
Private Const MAX_COLS As Long = 26
Private Const MAX_CELL As Long = 160
Private Const MAX_SHEETS As Long = 8
Private Const MAX_SEL_CELLS As Long = 300

'--- number/text conversion ---------------------------------------------------
' Str$() is the only VBA number formatter that ignores the regional decimal
' separator, which matters because the bridge always wants "1234.5", never
' "1234,5".
Private Function NumToText(ByVal d As Double) As String
    Dim s As String
    s = Trim$(Str$(d))
    If s = "-0" Then s = "0"
    NumToText = s
End Function

Public Function CellToText(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Then
        CellToText = "#ERR"
        Exit Function
    End If
    If IsEmpty(v) Or IsNull(v) Then
        CellToText = ""
        Exit Function
    End If
    Select Case VarType(v)
        Case vbDate
            s = Format$(v, "yyyy-mm-dd")
        Case vbBoolean
            s = IIf(v, "TRUE", "FALSE")
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal
            s = NumToText(CDbl(v))
        Case Else
            s = CStr(v)
    End Select
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    If Len(s) > MAX_CELL Then s = Left$(s, MAX_CELL) & "..."
    CellToText = s
End Function

'--- range -> grid record -----------------------------------------------------
Private Function GridRecord(ByVal ws As Worksheet) As String
    Dim ur As Range, rows As Long, cols As Long
    Dim vals As Variant, r As Long, c As Long, i As Long
    Dim flat() As String, n As Long

    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then Exit Function

    ' A whole-column selection makes UsedRange enormous; skip rather than hang.
    If ur.Cells.CountLarge > 250000 Then Exit Function
    If Application.WorksheetFunction.CountA(ur) = 0 Then Exit Function

    rows = ur.Rows.Count
    cols = ur.Columns.Count
    If rows > MAX_ROWS Then rows = MAX_ROWS
    If cols > MAX_COLS Then cols = MAX_COLS
    If rows < 1 Or cols < 1 Then Exit Function

    ReDim flat(0 To rows * cols - 1)
    n = 0
    If rows * cols = 1 Then
        ' a 1x1 range hands back a scalar, not a 2D array
        flat(0) = CellToText(ur.Cells(1, 1).Value2)
    Else
        vals = ur.Resize(rows, cols).Value2
        If Not IsArray(vals) Then Exit Function
        For r = 1 To rows
            For c = 1 To cols
                flat(n) = CellToText(vals(r, c))
                n = n + 1
            Next c
        Next r
    End If

    GridRecord = Rec("GRID", ws.Name, ur.Cells(1, 1).Address(False, False), _
                     CStr(rows), CStr(cols), Join(flat, SEP))
End Function

'--- selection ----------------------------------------------------------------
Private Function SelectionRecord() As String
    Dim rng As Range, ws As Worksheet
    Dim rows As Long, cols As Long, vals As Variant, flat() As String
    Dim r As Long, c As Long, n As Long

    If TypeName(Selection) <> "Range" Then Exit Function
    On Error Resume Next
    Set rng = Selection
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    If rng.Cells.CountLarge > MAX_SEL_CELLS Then Exit Function

    Set ws = rng.Worksheet
    rows = rng.Rows.Count
    cols = rng.Columns.Count
    If rows * cols > MAX_SEL_CELLS Then Exit Function
    If rows * cols = 0 Then Exit Function

    vals = rng.Value2
    ReDim flat(0 To rows * cols - 1)
    n = 0
    If rows = 1 And cols = 1 Then
        flat(0) = CellToText(vals)
    Else
        For r = 1 To rows
            For c = 1 To cols
                flat(n) = CellToText(vals(r, c))
                n = n + 1
            Next c
        Next r
    End If

    SelectionRecord = Rec("SELECTION", ws.Name, rng.Address(False, False), _
                          CStr(rows), CStr(cols), Join(flat, SEP))
End Function

'--- the whole request body ---------------------------------------------------
Public Function BuildRequest(ByVal userText As String, ByVal readonlyMode As Boolean, _
                             ByVal modelName As String, _
                             ByRef histRole() As String, ByRef histText() As String, _
                             ByVal histCount As Long) As String
    Dim wb As Workbook, ws As Worksheet, used As Range
    Dim sb As String, i As Long, n As Long

    On Error Resume Next
    Set wb = ActiveWorkbook
    On Error GoTo 0
    If wb Is Nothing Then
        BuildRequest = Rec("REPLY", "No workbook is open.") & vbLf & "END"
        Exit Function
    End If

    modelName = Trim$(modelName)
    If Len(modelName) = 0 Then modelName = DEFAULT_MODEL

    sb = Rec("VERSION", "1") & vbLf
    If readonlyMode Then sb = sb & Rec("READONLY", "1") & vbLf
    sb = sb & Rec("MODEL", modelName) & vbLf
    sb = sb & Rec("USER", userText) & vbLf
    sb = sb & Rec("ACTIVESHEET", ActiveSheet.Name) & vbLf

    ' sheet inventory first, so the model knows the shape before the grids.
    ' Iterate Sheets (not Worksheets) and filter: chart sheets would otherwise
    ' blow up the Worksheet type.
    n = 0
    For Each ws In wb.Worksheets
        n = n + 1
        If n > MAX_SHEETS Then Exit For
        Set used = Nothing
        On Error Resume Next
        Set used = ws.UsedRange
        On Error GoTo 0
        If used Is Nothing Then
            sb = sb & Rec("SHEET", ws.Name, "", "0", "0") & vbLf
        Else
            sb = sb & Rec("SHEET", ws.Name, used.Address(False, False), _
                          CStr(used.Rows.Count), CStr(used.Columns.Count)) & vbLf
        End If
    Next ws

    sb = sb & SelectionRecord() & vbLf

    n = 0
    For Each ws In wb.Worksheets
        n = n + 1
        If n > MAX_SHEETS Then Exit For
        sb = sb & GridRecord(ws) & vbLf
    Next ws

    For i = 1 To histCount
        sb = sb & Rec("HISTORY", histRole(i), histText(i)) & vbLf
    Next i

    sb = sb & "END"
    BuildRequest = sb
End Function
