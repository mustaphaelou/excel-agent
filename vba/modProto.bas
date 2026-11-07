Option Explicit
'==============================================================================
' modProto - wire format shared with bridge.py
'
' Records are separated by CRLF, fields by Chr(31) (unit separator).
' Inside a field:  \\ -> \   \n -> LF   \r -> CR   \t -> TAB   \s -> Chr(31)
' Keep this in lockstep with esc()/unesc() in bridge.py.
'==============================================================================

Public Function SEP() As String
    SEP = Chr$(31)
End Function

'--- field escaping -----------------------------------------------------------
Public Function Esc(ByVal s As String) As String
    Dim r As String
    r = Replace(s, "\", "\\")
    r = Replace(r, vbCr, "\r")
    r = Replace(r, vbLf, "\n")
    r = Replace(r, vbTab, "\t")
    r = Replace(r, SEP, "\s")
    Esc = r
End Function

Public Function Unesc(ByVal s As String) As String
    Dim i As Long, n As Long, out As String, c As String, nx As String
    n = Len(s)
    i = 1
    Do While i <= n
        c = Mid$(s, i, 1)
        If c = "\" And i < n Then
            nx = Mid$(s, i + 1, 1)
            Select Case nx
                Case "\": out = out & "\"
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case "s": out = out & SEP
                Case Else: out = out & nx
            End Select
            i = i + 2
        Else
            out = out & c
            i = i + 1
        End If
    Loop
    Unesc = out
End Function

'--- record helpers -----------------------------------------------------------
Public Function Rec(ParamArray fields() As Variant) As String
    Dim i As Long, out As String
    For i = LBound(fields) To UBound(fields)
        If i > LBound(fields) Then out = out & SEP
        out = out & Esc(CStr(fields(i)))
    Next
    Rec = out
End Function

Public Function RecParts(ByVal recLine As String) As String()
    Dim parts() As String, i As Long
    parts = Split(recLine, SEP)
    For i = LBound(parts) To UBound(parts)
        parts(i) = Unesc(parts(i))
    Next
    RecParts = parts
End Function

'--- response stream parser ---------------------------------------------------
' Splits the bridge answer into records. Each record is a String() of fields.
Public Function ParseResponse(ByVal body As String) As Collection
    Dim col As New Collection
    Dim lines() As String, i As Long, ln As String, parts() As String
    body = Replace(body, vbCrLf, vbLf)
    body = Replace(body, vbCr, vbLf)
    lines = Split(body, vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = lines(i)
        If Len(Trim$(ln)) > 0 Then
            parts = RecParts(ln)
            If UBound(parts) >= 0 Then col.Add parts
        End If
    Next
    Set ParseResponse = col
End Function

'--- loose change-of-record wrapper used by the form -------------------------
Public Function KindOf(ByRef parts() As String) As String
    If UBound(parts) >= 0 Then KindOf = UCase$(Trim$(parts(0)))
End Function

Public Function FieldOr(ByRef parts() As String, ByVal idx As Long, ByVal dflt As String) As String
    If UBound(parts) >= idx Then FieldOr = parts(idx) Else FieldOr = dflt
End Function
