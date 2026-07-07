Attribute VB_Name = "modFormHelpers"
Option Explicit

'------------------------------------------------------------------------------
' Module : modFormHelpers
' Purpose : Shared UserForm validation helpers.
' Notes : Form layout is designed in the VBA UserForm designer; only visibility
' and enabled state are changed in code.
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
' Purpose : Parses a trimmed date string into a Date value.
' Output : True when inputValue is a valid date.
'------------------------------------------------------------------------------
Public Function TryGetDate(ByVal inputValue As String, ByRef outputDate As Date) As Boolean

    inputValue = Trim$(inputValue)

    If Len(inputValue) = 0 Then
        TryGetDate = False
        Exit Function
    End If

    If IsDate(inputValue) Then
        outputDate = CDate(inputValue)
        TryGetDate = True
    Else
        TryGetDate = False
    End If

End Function
