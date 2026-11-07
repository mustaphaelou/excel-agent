Option Explicit
'==============================================================================
' modApply - executes the ACTION records the bridge returns.
'
' Wire shape of every record:
'   ACTION | kind | human label | arg1 | arg2 | ...
'
' Every action is validated again here. The model is not trusted: a sheet that
' does not exist, a bad range address, or a nonsense property fails as a
' logged message instead of an error dialog.
'
' VBA edits are not in Excel's undo stack, which is why the form keeps a
' confirm step before calling in here.
'==============================================================================

'--- lookups ------------------------------------------------------------------
Public Function GetSheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, Trim$(sheetName), vbTextCompare) = 0 Then
            Set GetSheetByName = ws
            Exit Function
        End If
    Next ws
End Function

Public Function ColorFromHex(ByVal s As String) As Long
    Dim r As Long, g As Long, b As Long
    s = Replace(Trim$(s), "#", "")
    If Len(s) <> 6 Then Exit Function
    r = CLng("&H" & Mid$(s, 1, 2))
    g = CLng("&H" & Mid$(s, 3, 2))
    b = CLng("&H" & Mid$(s, 5, 2))
    ColorFromHex = RGB(r, g, b)
End Function

Private Function CDblSafe(ByVal s As String) As Double
    Dim t As String
    t = Trim$(s)
    t = Replace(t, ",", ".")
    If IsPlainNumber(t) Then CDblSafe = Val(t) Else CDblSafe = 0
End Function

Private Function IsTrue(ByVal s As String) As Boolean
    IsTrue = (s = "1" Or UCase$(Trim$(s)) = "TRUE")
End Function

'--- value coercion -----------------------------------------------------------
' The model returns everything as JSON text. "12.5" must land as a number, not
' as the string "12.5", or SUM() silently counts it as zero.
Public Function IsPlainNumber(ByVal t As String) As Boolean
    Dim i As Long, c As String, digits As Long, dots As Long
    If Len(t) = 0 Then Exit Function
    If Len(t) > 1 And (Left$(t, 1) = "-" Or Left$(t, 1) = "+") Then t = Mid$(t, 2)
    For i = 1 To Len(t)
        c = Mid$(t, i, 1)
        If c >= "0" And c <= "9" Then
            digits = digits + 1
        ElseIf c = "." Then
            dots = dots + 1
            If dots > 1 Then Exit Function
        Else
            Exit Function
        End If
    Next i
    IsPlainNumber = (digits > 0)
End Function

Private Function CellValue(ByVal s As String) As Variant
    Dim t As String
    t = Trim$(s)
    If Len(t) = 0 Then
        CellValue = Empty
    ElseIf IsPlainNumber(t) Then
        CellValue = Val(t)
    Else
        CellValue = s
    End If
End Function

'--- dispatch -----------------------------------------------------------------
Public Function ApplyOne(ByRef p() As String, ByRef wb As Workbook) As String
    Dim kind As String
    If UBound(p) < 3 Then
        ApplyOne = "Malformed action record skipped."
        Exit Function
    End If
    kind = UCase$(Trim$(p(1)))

    On Error GoTo Failed
    Select Case kind
        Case "SET_VALUES":   ApplyOne = DoSetValues(p, wb)
        Case "SET_FORMULA":  ApplyOne = DoSetFormula(p, wb)
        Case "FORMAT":       ApplyOne = DoFormat(p, wb)
        Case "AUTOFIT":      ApplyOne = DoAutofit(p, wb)
        Case "CREATE_SHEET": ApplyOne = DoCreateSheet(p, wb)
        Case "RENAME_SHEET": ApplyOne = DoRenameSheet(p, wb)
        Case "DELETE_SHEET": ApplyOne = DoDeleteSheet(p, wb)
        Case "SORT":         ApplyOne = DoSort(p, wb)
        Case "CLEAR":        ApplyOne = DoClear(p, wb)
        Case "SELECT_CELL":  ApplyOne = DoSelect(p, wb)
        Case Else
            ApplyOne = "Unknown action '" & kind & "' skipped."
    End Select
    Exit Function
Failed:
    ApplyOne = "Failed (" & kind & "): " & Err.Description
End Function

'--- set_values ---------------------------------------------------------------
' p(3)=sheet p(4)=anchor p(5)=rows p(6)=cols p(7..)=values (row-major)
Private Function DoSetValues(ByRef p() As String, ByRef wb As Workbook) As String
    Dim ws As Worksheet, anchor As Range, out As Range
    Dim rows As Long, cols As Long, i As Long, r As Long, c As Long
    Dim arr As Variant, formulas As Collection
    Dim txt As String

    Set ws = GetSheetByName(wb, p(3))
    If ws Is Nothing Then DoSetValues = "No sheet '" & p(3) & "'.": Exit Function

    rows = CLng(p(5))
    cols = CLng(p(6))
    If rows < 1 Or cols < 1 Then DoSetValues = "Empty block skipped.": Exit Function
    If UBound(p) < 6 + rows * cols Then
        DoSetValues = "Block size mismatch on " & ws.Name & "!" & p(4) & "."
        Exit Function
    End If

    Set anchor = ws.Range(p(4))
    Set out = anchor.Resize(rows, cols)

    ReDim arr(1 To rows, 1 To cols)
    Set formulas = New Collection
    i = 7
    For r = 1 To rows
        For c = 1 To cols
            txt = p(i)
            If Len(txt) > 0 And Left$(txt, 1) = "=" Then
                arr(r, c) = Empty
                formulas.Add Array(r, c, txt)
            Else
                arr(r, c) = CellValue(txt)
            End If
            i = i + 1
        Next c
    Next r

    out.Value2 = arr

    Dim f As Variant
    For Each f In formulas
        out.Cells(f(0), f(1)).Formula = f(2)
    Next f
End Function

'--- set_formula -------------------------------------------------------------- p(3)=sheet p(4)=range p(5)=formula
Private Function DoSetFormula(ByRef p() As String, ByRef wb As Workbook) As String
    Dim ws As Worksheet, rng As Range
    Set ws = GetSheetByName(wb, p(3))
    If ws Is Nothing Then DoSetFormula = "No sheet '" & p(3) & "'.": Exit Function
    Set rng = ws.Range(p(4))
    rng.Formula = p(5)
End Function

'--- format ------------------------------------------------------------------- p(3)=sheet p(4)=range p(5)=prop p(6)=value
Private Function DoFormat(ByRef p() As String, ByRef wb As Workbook) As String
    Dim ws As Worksheet, rng As Range, prop As String, sVal As String
    Set ws = GetSheetByName(wb, p(3))
    If ws Is Nothing Then DoFormat = "No sheet '" & p(3) & "'.": Exit Function
    Set rng = ws.Range(p(4))
    prop = LCase$(Trim$(p(5)))
    sVal = p(6)

    Select Case prop
        Case "bold":          rng.Font.Bold = IsTrue(sVal)
        Case "italic":        rng.Font.Italic = IsTrue(sVal)
        Case "wrap":          rng.WrapText = IsTrue(sVal)
        Case "borders"
            If IsTrue(sVal) Then
                rng.Borders.LineStyle = 1              ' xlContinuous
            Else
                rng.Borders.LineStyle = -4142          ' xlLineStyleNone
            End If
        Case "fill"
            rng.Interior.Pattern = 1                   ' xlSolid
            rng.Interior.Color = ColorFromHex(sVal)
        Case "font_color"
            rng.Font.Color = ColorFromHex(sVal)
        Case "number_format"
            rng.NumberFormat = sVal
        Case "column_width"
            rng.ColumnWidth = CDblSafe(sVal)
        Case "row_height"
            rng.RowHeight = CDblSafe(sVal)
        Case "h_align"
            Select Case LCase$(sVal)
                Case "center":  rng.HorizontalAlignment = -4108
                Case "left":    rng.HorizontalAlignment = -4131
                Case "right":   rng.HorizontalAlignment = -4152
                Case "general": rng.HorizontalAlignment = 1
                Case Else:      DoFormat = "Unknown alignment '" & sVal & "'."
            End Select
        Case Else
            DoFormat = "Unknown format property '" & prop & "'."
    End Select
End Function

'--- autofit ------------------------------------------------------------------ p(3)=sheet p(4)=range p(5)=columns|rows
Private Function DoAutofit(ByRef p() As String, ByRef wb As Workbook) As String
    Dim ws As Worksheet, rng As Range
    Set ws = GetSheetByName(wb, p(3))
    If ws Is Nothing Then DoAutofit = "No sheet '" & p(3) & "'.": Exit Function
    Set rng = ws.Range(p(4))
    If LCase$(Trim$(p(5))) = "rows" Then
        rng.EntireRow.AutoFit
    Else
        rng.EntireColumn.AutoFit
    End If
End Function

'--- structure ----------------------------------------------------------------
Private Function DoCreateSheet(ByRef p() As String, ByRef wb As Workbook) As String
    Dim sheetName As String, ws As Worksheet
    sheetName = Trim$(p(3))
    If Len(sheetName) = 0 Then DoCreateSheet = "Empty sheet name.": Exit Function
    If Not (GetSheetByName(wb, sheetName) Is Nothing) Then
        DoCreateSheet = "Sheet '" & sheetName & "' already exists; not created."
        Exit Function
    End If
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    On Error Resume Next
    ws.Name = sheetName
    If Err.Number <> 0 Then
        DoCreateSheet = "Could not rename sheet '" & sheetName & "': " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function DoRenameSheet(ByRef p() As String, ByRef wb As Workbook) As String
    Dim ws As Worksheet
    Set ws = GetSheetByName(wb, p(3))
    If ws Is Nothing Then DoRenameSheet = "No sheet '" & p(3) & "'.": Exit Function
    If Not (GetSheetByName(wb, p(4)) Is Nothing) Then
        DoRenameSheet = "A sheet named '" & p(4) & "' already exists."
        Exit Function
    End If
    ws.Name = p(4)
End Function

Private Function DoDeleteSheet(ByRef p() As String, ByRef wb As Workbook) As String
    Dim ws As Worksheet
    If wb.Worksheets.Count <= 1 Then
        DoDeleteSheet = "Refused: that is the only sheet in the workbook."
        Exit Function
    End If
    Set ws = GetSheetByName(wb, p(3))
    If ws Is Nothing Then DoDeleteSheet = "No sheet '" & p(3) & "'.": Exit Function
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
End Function

'--- sort --------------------------------------------------------------------- p(3)=sheet p(4)=range p(5)=keycol p(6)=order p(7)=hasheader
Private Function DoSort(ByRef p() As String, ByRef wb As Workbook) As String
    Dim ws As Worksheet, rng As Range, keyCol As Long, ordr As Long, hdr As Long
    Set ws = GetSheetByName(wb, p(3))
    If ws Is Nothing Then DoSort = "No sheet '" & p(3) & "'.": Exit Function
    Set rng = ws.Range(p(4))

    keyCol = CLng(p(5))
    If keyCol < 1 Or keyCol > rng.Columns.Count Then
        DoSort = "Sort key column " & keyCol & " is outside " & p(4) & "."
        Exit Function
    End If
    ordr = IIf(LCase$(Trim$(p(6))) = "desc", 2, 1)     ' xlDescending / xlAscending
    hdr = IIf(p(7) = "1", 1, 2)                        ' xlYes / xlNo

    rng.Sort Key1:=rng.Columns(keyCol), Order1:=ordr, Header:=hdr
End Function

'--- clear -------------------------------------------------------------------- p(3)=sheet p(4)=range p(5)=contents|formats|all
Private Function DoClear(ByRef p() As String, ByRef wb As Workbook) As String
    Dim ws As Worksheet, rng As Range, what As String
    Set ws = GetSheetByName(wb, p(3))
    If ws Is Nothing Then DoClear = "No sheet '" & p(3) & "'.": Exit Function
    Set rng = ws.Range(p(4))
    what = LCase$(Trim$(p(5)))
    Select Case what
        Case "formats": rng.ClearFormats
        Case "all":     rng.Clear
        Case Else:      rng.ClearContents
    End Select
End Function

'--- select ------------------------------------------------------------------- p(3)=sheet p(4)=address
Private Function DoSelect(ByRef p() As String, ByRef wb As Workbook) As String
    Dim ws As Worksheet
    Set ws = GetSheetByName(wb, p(3))
    If ws Is Nothing Then DoSelect = "No sheet '" & p(3) & "'.": Exit Function
    Application.Goto ws.Range(p(4)), True
End Function
