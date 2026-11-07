Option Explicit
'==============================================================================
' TurnResult - the outcome of one Turn.Submit call.
'
' The panel reads this and renders it. It never sees a raw ACTION record:
' Edits holds display-ready strings ("[kind] label") while Turn keeps the
' records themselves for Apply.
'==============================================================================

Public Ok As Boolean
Public Reply As String
Public ErrorText As String
Public Status As String
Public Notes As Collection
Public Edits As Collection

Private Sub Class_Initialize()
    Set Notes = New Collection
    Set Edits = New Collection
End Sub
