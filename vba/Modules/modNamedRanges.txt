Attribute VB_Name = "modNamedRanges"
Option Explicit

'------------------------------------------------------------------------------
' Module : modNamedRanges
' Purpose : Provides safe access to workbook-level named ranges.
'
' Notes:
' - These functions expect workbook-level names in ThisWorkbook.
' - If a named range is missing or contains invalid data, a clear VBA error is
' raised so the calling procedure can show a useful message to the user.
'------------------------------------------------------------------------------

Private Const ERR_BASE_NAMED_RANGES As Long = vbObjectError + 2000


'------------------------------------------------------------------------------
' Purpose : Returns the value from a workbook-level named range.
' Input : rangeName - the workbook name to read, e.g. "ForecastStart".
' Output : The value stored in the named range.
' Raises : Error if the name is missing, invalid, or does not refer to a range.
'------------------------------------------------------------------------------
Public Function GetNamedRangeValue(ByVal rangeName As String) As Variant

    On Error GoTo ErrHandler

    GetNamedRangeValue = ThisWorkbook.Names(rangeName).RefersToRange.Value
    Exit Function

ErrHandler:
    Err.Raise ERR_BASE_NAMED_RANGES + 1, _
              "GetNamedRangeValue", _
              "The named range '" & rangeName & "' is missing or invalid."

End Function


'------------------------------------------------------------------------------
' Purpose : Returns the worksheet range behind a workbook-level named range.
' Input : rangeName - the workbook name to read.
' Output : Range object referred to by the name.
' Raises : Error if the name is missing, invalid, or does not refer to a range.
'------------------------------------------------------------------------------
Public Function GetNamedRange(ByVal rangeName As String) As Range

    On Error GoTo ErrHandler

    Set GetNamedRange = ThisWorkbook.Names(rangeName).RefersToRange
    Exit Function

ErrHandler:
    Err.Raise ERR_BASE_NAMED_RANGES + 2, _
              "GetNamedRange", _
              "The named range '" & rangeName & "' is missing or invalid."

End Function


'------------------------------------------------------------------------------
' Purpose : Reads a named range that must contain a valid date.
' Input : rangeName - the workbook name to read.
' Output : Date value from the named range.
'------------------------------------------------------------------------------
Public Function GetNamedDate(ByVal rangeName As String) As Date

    Dim rawValue As Variant
    rawValue = GetNamedRangeValue(rangeName)

    If Not IsDate(rawValue) Then
        Err.Raise ERR_BASE_NAMED_RANGES + 3, _
                  "GetNamedDate", _
                  "The named range '" & rangeName & "' must contain a valid date."
    End If

    GetNamedDate = CDate(rawValue)

End Function


'------------------------------------------------------------------------------
' Purpose : Reads a named range that must contain a whole number.
' Input : rangeName - the workbook name to read.
' Output : Long value from the named range.
'------------------------------------------------------------------------------
Public Function GetNamedLong(ByVal rangeName As String) As Long

    Dim rawValue As Variant
    rawValue = GetNamedRangeValue(rangeName)

    If Not IsNumeric(rawValue) Then
        Err.Raise ERR_BASE_NAMED_RANGES + 4, _
                  "GetNamedLong", _
                  "The named range '" & rangeName & "' must contain a whole number."
    End If

    If CDbl(rawValue) <> Fix(CDbl(rawValue)) Then
        Err.Raise ERR_BASE_NAMED_RANGES + 5, _
                  "GetNamedLong", _
                  "The named range '" & rangeName & "' must contain a whole number, not a decimal."
    End If

    GetNamedLong = CLng(rawValue)

End Function


'------------------------------------------------------------------------------
' Purpose : Reads a named range that must contain a number.
' Input : rangeName - the workbook name to read.
' Output : Double value from the named range.
'------------------------------------------------------------------------------
Public Function GetNamedDouble(ByVal rangeName As String) As Double

    Dim rawValue As Variant
    rawValue = GetNamedRangeValue(rangeName)

    If Not IsNumeric(rawValue) Then
        Err.Raise ERR_BASE_NAMED_RANGES + 6, _
                  "GetNamedDouble", _
                  "The named range '" & rangeName & "' must contain a number."
    End If

    GetNamedDouble = CDbl(rawValue)

End Function


'------------------------------------------------------------------------------
' Purpose : Reads a named range that represents a Yes/No value.
' Input : rangeName - the workbook name to read.
' Output : True for Y/YES/TRUE/1, False for N/NO/FALSE/0/blank.
'------------------------------------------------------------------------------
Public Function GetNamedYesNo(ByVal rangeName As String) As Boolean

    Dim rawValue As String
    rawValue = UCase$(Trim$(CStr(GetNamedRangeValue(rangeName))))

    Select Case rawValue

        Case "Y", "YES", "TRUE", "1"
            GetNamedYesNo = True

        Case "N", "NO", "FALSE", "0", vbNullString
            GetNamedYesNo = False

        Case Else
            Err.Raise ERR_BASE_NAMED_RANGES + 7, _
                      "GetNamedYesNo", _
                      "The named range '" & rangeName & "' must contain Y or N."

    End Select

End Function


'------------------------------------------------------------------------------
' Purpose : Reads a named range as trimmed text.
' Input : rangeName - the workbook name to read.
' Output : Trimmed string value.
'------------------------------------------------------------------------------
Public Function GetNamedText(ByVal rangeName As String) As String

    GetNamedText = Trim$(CStr(GetNamedRangeValue(rangeName)))

End Function


