Attribute VB_Name = "modPlannerRates"
Option Explicit

'------------------------------------------------------------------------------
' Module : modPlannerRates
' Purpose : Loads and stores planned weekly flying rates.
'
' Current planner version:
' Supports a fixed 3 flying week / 2 maintenance week cycle.
'
' Rate storage:
' Rates are read from workbook-level named ranges, not fixed Settings cells.
' This allows the Settings sheet layout to change without breaking refresh.
'
' Required named ranges:
' FlyingWeek1HHRate
' FlyingWeek1E1Rate
' FlyingWeek1E2Rate
'
' FlyingWeek2HHRate
' FlyingWeek2E1Rate
' FlyingWeek2E2Rate
'
' FlyingWeek3HHRate
' FlyingWeek3E1Rate
' FlyingWeek3E2Rate
'
' Notes:
' Settings cells store rates as Excel durations formatted [h]:mm.
' VBA calculations use decimal hours, so values are multiplied by 24.
'
' Future improvement:
' When variable flying weeks are supported, replace FIXED_FLYING_WEEKS
' and fixed-size arrays with dynamic arrays sized from PlannerSettings.
'------------------------------------------------------------------------------

Private Const FIXED_FLYING_WEEKS As Long = 3

Private Const RATE_TYPE_HH As String = "HH"
Private Const RATE_TYPE_E1 As String = "E1"
Private Const RATE_TYPE_E2 As String = "E2"

Private Const ERR_BASE_PLANNER_RATES As Long = vbObjectError + 2200


'------------------------------------------------------------------------------
' Type : PlannerRates
' Purpose : Stores flying-rate assumptions for each flying week in the cycle.
'
' Current limitation:
' Fixed to 3 flying weeks.
'------------------------------------------------------------------------------
Public Type plannerRates
    HHRates(1 To FIXED_FLYING_WEEKS) As Double
    E1Rates(1 To FIXED_FLYING_WEEKS) As Double
    E2Rates(1 To FIXED_FLYING_WEEKS) As Double
End Type


'------------------------------------------------------------------------------
' Purpose : Loads planned flying rates from named ranges.
' Output : PlannerRates value containing HH, E1, and E2 rates.
' Raises : Error if any required rate named range is missing or invalid.
'------------------------------------------------------------------------------
Public Function LoadPlannerRates() As plannerRates

    Dim rates As plannerRates
    Dim flyingWeek As Long

    For flyingWeek = 1 To FIXED_FLYING_WEEKS

        rates.HHRates(flyingWeek) = GetRateValueByName(flyingWeek, RATE_TYPE_HH)
        rates.E1Rates(flyingWeek) = GetRateValueByName(flyingWeek, RATE_TYPE_E1)
        rates.E2Rates(flyingWeek) = GetRateValueByName(flyingWeek, RATE_TYPE_E2)

    Next flyingWeek

    LoadPlannerRates = rates

End Function


'------------------------------------------------------------------------------
' Purpose : Returns the decimal-hour rate for one flying week and counter type.
' Input : flyingWeek - flying week number, currently 1 to 3.
' rateType - HH, E1, or E2.
' Output : Decimal hours for calculation use.
'------------------------------------------------------------------------------
Private Function GetRateValueByName(ByVal flyingWeek As Long, _
                                    ByVal rateType As String) As Double

    Dim rangeName As String
    rangeName = BuildRateRangeName(flyingWeek, rateType)

    Dim displayName As String
    displayName = "Flying Week " & flyingWeek & " " & rateType

    GetRateValueByName = GetDurationNamedRangeAsDecimalHours(rangeName, displayName)

End Function


'------------------------------------------------------------------------------
' Purpose : Builds the named range used for one flying week/rate type.
' Example : FlyingWeek2E2Rate
'------------------------------------------------------------------------------
Private Function BuildRateRangeName(ByVal flyingWeek As Long, _
                                    ByVal rateType As String) As String

    ValidateFlyingWeekNumber flyingWeek
    ValidateRateType rateType

    BuildRateRangeName = "FlyingWeek" & CStr(flyingWeek) & rateType & "Rate"

End Function


'------------------------------------------------------------------------------
' Purpose : Reads one [h]:mm duration named range and returns decimal hours.
' Notes : Excel stores durations as fractions of a day.
' Example: 10:00 is stored as 10 / 24.
'------------------------------------------------------------------------------
Private Function GetDurationNamedRangeAsDecimalHours(ByVal rangeName As String, _
                                                     ByVal displayName As String) As Double

    Dim rateCell As Range
    Set rateCell = GetSingleCellNamedRange(rangeName, displayName)

    If Not IsNumeric(rateCell.Value) Then
        Err.Raise ERR_BASE_PLANNER_RATES + 1, _
                  "GetDurationNamedRangeAsDecimalHours", _
                  displayName & " must contain a numeric [h]:mm duration." & vbCrLf & _
                  "Named range: " & rangeName & vbCrLf & _
                  "Cell: " & rateCell.Address(External:=True) & vbCrLf & _
                  "Current value: [" & CStr(rateCell.Value) & "]"
    End If

    If CDbl(rateCell.Value) < 0 Then
        Err.Raise ERR_BASE_PLANNER_RATES + 2, _
                  "GetDurationNamedRangeAsDecimalHours", _
                  displayName & " cannot be negative." & vbCrLf & _
                  "Named range: " & rangeName & vbCrLf & _
                  "Cell: " & rateCell.Address(External:=True)
    End If

    GetDurationNamedRangeAsDecimalHours = CDbl(rateCell.Value) * 24#

End Function


'------------------------------------------------------------------------------
' Purpose : Returns a named range and confirms it refers to one cell.
'------------------------------------------------------------------------------
Private Function GetSingleCellNamedRange(ByVal rangeName As String, _
                                         ByVal displayName As String) As Range

    On Error GoTo ErrHandler

    Dim targetRange As Range
    Set targetRange = ThisWorkbook.Names(rangeName).RefersToRange

    If targetRange.Cells.CountLarge <> 1 Then
        Err.Raise ERR_BASE_PLANNER_RATES + 3, _
                  "GetSingleCellNamedRange", _
                  displayName & " named range must refer to one cell only." & vbCrLf & _
                  "Named range: " & rangeName & vbCrLf & _
                  "Current reference: " & targetRange.Address(External:=True)
    End If

    Set GetSingleCellNamedRange = targetRange
    Exit Function

ErrHandler:
    Err.Raise ERR_BASE_PLANNER_RATES + 4, _
              "GetSingleCellNamedRange", _
              "Could not read named range for " & displayName & "." & vbCrLf & _
              "Named range: " & rangeName & vbCrLf & _
              "Error " & Err.Number & ": " & Err.Description

End Function


'------------------------------------------------------------------------------
' Purpose : Validates the currently supported flying week number.
'------------------------------------------------------------------------------
Private Sub ValidateFlyingWeekNumber(ByVal flyingWeek As Long)

    If flyingWeek < 1 Or flyingWeek > FIXED_FLYING_WEEKS Then
        Err.Raise ERR_BASE_PLANNER_RATES + 5, _
                  "ValidateFlyingWeekNumber", _
                  "Flying week " & flyingWeek & " is outside the supported range." & vbCrLf & _
                  "Current planner version supports flying weeks 1 to " & FIXED_FLYING_WEEKS & "."
    End If

End Sub


'------------------------------------------------------------------------------
' Purpose : Validates supported rate type.
'------------------------------------------------------------------------------
Private Sub ValidateRateType(ByVal rateType As String)

    Select Case UCase$(Trim$(rateType))

        Case RATE_TYPE_HH, RATE_TYPE_E1, RATE_TYPE_E2
            ' Valid.

        Case Else
            Err.Raise ERR_BASE_PLANNER_RATES + 6, _
                      "ValidateRateType", _
                      "Unsupported rate type: " & rateType

    End Select

End Sub


