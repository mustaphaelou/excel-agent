Option Explicit
'==============================================================================
' modWindow - Win32 bits the MSForms runtime does not expose.
'
' A VBA UserForm is created with a fixed-border window style. Adding
' WS_THICKFRAME gives it real resize edges and a maximise button; from then on
' the user can drag the panel around like any other window.
'
' Declarations are kept on single lines and the conditional compilation is
' flat (no nested #If): VBA rejects this construct when a Declare is continued
' across lines inside nested conditionals ("Compile error: Syntax error").
'==============================================================================

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLongA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongA Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
#ElseIf VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
#Else
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const WS_MINIMIZEBOX As Long = &H20000

'--- locate a UserForm's window -------------------------------------------------
' MSForms windows use the class name "ThunderDFrame". If that ever changes we
' fall back to a title match, which is unique enough for our caption.
Public Function FormHwnd(ByVal caption As String) As LongPtr
    Dim h As LongPtr
    On Error Resume Next
    h = FindWindowA("ThunderDFrame", caption)
    If h = 0 Then h = FindWindowA(vbNullString, caption)
    FormHwnd = h
End Function

'--- give a form resize edges + maximise button --------------------------------
Public Function MakeResizable(ByVal caption As String) As Boolean
    Dim h As LongPtr, style As LongPtr
    On Error GoTo Failed

    h = FormHwnd(caption)
    If h = 0 Then Exit Function

    style = GetWindowLongA(h, GWL_STYLE)
    style = style Or WS_THICKFRAME Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX
    Call SetWindowLongA(h, GWL_STYLE, style)
    Call DrawMenuBar(h)

    MakeResizable = True
    Exit Function
Failed:
    MakeResizable = False
End Function
