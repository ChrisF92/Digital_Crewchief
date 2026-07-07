Attribute VB_Name = "modAircraftStatus"
Option Explicit

'------------------------------------------------------------------------------
' Module : modAircraftStatus
' Purpose : Calculates the current dashboard status for an aircraft.
'------------------------------------------------------------------------------

Public Type AircraftCurrentStatus
    RotationText As String
    currentWeekIndex As Long
End Type


'------------------------------------------------------------------------------
' Purpose : Calculates the current visible status for one aircraft.
' Input : settings - validated planner settings.
' weekStartDates - visible planner week start dates.
' weekCyclePosition - weekly cycle position array for this aircraft.
' isDownMaintenance - True when aircraft is in long-term maintenance.
' rtsFlyStart - first Monday after ETBOL.
' maintenanceReason - optional down-maintenance reason.
' Output : AircraftCurrentStatus with display text and current week index.
'------------------------------------------------------------------------------
Public Function GetAircraftCurrentStatus(ByRef settings As PlannerSettings, _
                                         ByRef weekStartDates() As Date, _
                                         ByRef weekCyclePosition() As Long, _
                                         ByVal isDownMaintenance As Boolean, _
                                         ByVal rtsFlyStart As Date, _
                                         ByVal maintenanceReason As String) As AircraftCurrentStatus

    Dim status As AircraftCurrentStatus

    Dim currentWeekStart As Date
    currentWeekStart = DisplayWeekStart(Date, settings.displayStart)

    Dim currentWeekIndex As Long
    currentWeekIndex = FindDisplayWeekIndex( _
        currentWeekStart, _
        weekStartDates, _
        settings.displayWeeksShown)

    If isDownMaintenance Then

        status = GetDownMaintenanceStatus( _
            settings, _
            weekCyclePosition, _
            currentWeekIndex, _
            rtsFlyStart, _
            maintenanceReason)

    Else

        status = GetInCycleAircraftStatus( _
            settings, _
            weekCyclePosition, _
            currentWeekIndex)

    End If

    GetAircraftCurrentStatus = status

End Function


'------------------------------------------------------------------------------
' Purpose : Calculates current status for a long-term down maintenance aircraft.
'------------------------------------------------------------------------------
Private Function GetDownMaintenanceStatus(ByRef settings As PlannerSettings, _
                                          ByRef weekCyclePosition() As Long, _
                                          ByVal currentWeekIndex As Long, _
                                          ByVal rtsFlyStart As Date, _
                                          ByVal maintenanceReason As String) As AircraftCurrentStatus

    Dim status As AircraftCurrentStatus

    If Date < rtsFlyStart Then

        If Len(Trim$(maintenanceReason)) > 0 Then
            status.RotationText = "Down Maint: " & Trim$(maintenanceReason)
        Else
            status.RotationText = "Down Maintenance"
        End If

        status.currentWeekIndex = 0

    ElseIf IsDisplayWeekIndexValid(currentWeekIndex, settings.displayWeeksShown) Then

        status.RotationText = FormatCyclePositionText( _
            weekCyclePosition(currentWeekIndex), _
            settings.flyingWeeks)

        status.currentWeekIndex = currentWeekIndex

    Else

        status.RotationText = "Post ETBOL"
        status.currentWeekIndex = currentWeekIndex

    End If

    GetDownMaintenanceStatus = status

End Function


'------------------------------------------------------------------------------
' Purpose : Calculates current status for a normal in-cycle aircraft.
'------------------------------------------------------------------------------
Private Function GetInCycleAircraftStatus(ByRef settings As PlannerSettings, _
                                          ByRef weekCyclePosition() As Long, _
                                          ByVal currentWeekIndex As Long) As AircraftCurrentStatus

    Dim status As AircraftCurrentStatus

    If Date < settings.forecastStart Then

        status.RotationText = "Pre-forecast"
        status.currentWeekIndex = 0

    ElseIf IsDisplayWeekIndexValid(currentWeekIndex, settings.displayWeeksShown) Then

        status.RotationText = FormatCyclePositionText( _
            weekCyclePosition(currentWeekIndex), _
            settings.flyingWeeks)

        status.currentWeekIndex = currentWeekIndex

    Else

        status.RotationText = "Outside forecast"
        status.currentWeekIndex = currentWeekIndex

    End If

    GetInCycleAircraftStatus = status

End Function


'------------------------------------------------------------------------------
' Purpose : Converts a cycle position into dashboard display text.
' Input : cyclePosition - position within the flying/maintenance cycle.
' flyingWeeks - number of flying weeks in the cycle.
' Output : Display text such as "Flying Wk 1" or "Maint Wk 2".
'------------------------------------------------------------------------------
Public Function FormatCyclePositionText(ByVal cyclePosition As Long, _
                                        ByVal flyingWeeks As Long) As String

    If cyclePosition <= flyingWeeks Then
        FormatCyclePositionText = "Flying Wk " & CStr(cyclePosition)
    Else
        FormatCyclePositionText = "Maint Wk " & CStr(cyclePosition - flyingWeeks)
    End If

End Function


