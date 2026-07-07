Attribute VB_Name = "modPlannerCalendar"
Option Explicit

'------------------------------------------------------------------------------
' Module : modPlannerCalendar
' Purpose : Builds and manages the planner display calendar.
'------------------------------------------------------------------------------

Private Const DAYS_PER_WEEK As Long = 7


'------------------------------------------------------------------------------
' Purpose : Builds an array of week start dates for the visible planner period.
' Input : displayStart - first week start date shown on the planner.
' displayWeeksShown - number of visible weeks.
' Output : 1-based Date array where each item is a week start date.
'------------------------------------------------------------------------------
Public Function BuildDisplayWeekStartDates(ByVal displayStart As Date, _
                                           ByVal displayWeeksShown As Long) As Date()

    If displayWeeksShown <= 0 Then
        Err.Raise vbObjectError + 2400, _
                  "BuildDisplayWeekStartDates", _
                  "DisplayWeeksShown must be greater than zero."
    End If

    Dim weekStartDates() As Date
    ReDim weekStartDates(1 To displayWeeksShown)

    Dim weekIndex As Long

    For weekIndex = 1 To displayWeeksShown
        weekStartDates(weekIndex) = displayStart + DAYS_PER_WEEK * (weekIndex - 1)
    Next weekIndex

    BuildDisplayWeekStartDates = weekStartDates

End Function


'------------------------------------------------------------------------------
' Purpose : Returns the planner index for a given week start date.
' Input : targetWeekStart - date to find.
' weekStartDates - 1-based array of display week start dates.
' displayWeeksShown - number of visible weeks.
' Output : Matching week index, or 0 if the week is not visible.
'------------------------------------------------------------------------------
Public Function FindDisplayWeekIndex(ByVal targetWeekStart As Date, _
                                     ByRef weekStartDates() As Date, _
                                     ByVal displayWeeksShown As Long) As Long

    Dim weekIndex As Long

    For weekIndex = 1 To displayWeeksShown
        If weekStartDates(weekIndex) = targetWeekStart Then
            FindDisplayWeekIndex = weekIndex
            Exit Function
        End If
    Next weekIndex

    FindDisplayWeekIndex = 0

End Function


'------------------------------------------------------------------------------
' Purpose : Returns True if a week index is inside the visible display range.
' Input : weekIndex - week index to check.
' displayWeeksShown - number of visible weeks.
'------------------------------------------------------------------------------
Public Function IsDisplayWeekIndexValid(ByVal weekIndex As Long, _
                                        ByVal displayWeeksShown As Long) As Boolean

    IsDisplayWeekIndexValid = (weekIndex >= 1 And weekIndex <= displayWeeksShown)

End Function


