Attribute VB_Name = "modMaintenanceWindow"
Option Explicit

'------------------------------------------------------------------------------
' Module : modMaintenanceWindow
' Purpose : Calculates the current or next maintenance window for an aircraft.
'------------------------------------------------------------------------------

Public Type MaintenanceWindowStatus
    startWeekIndex As Long
    endWeekIndex As Long
    daysToMaintenance As Variant
    IsCurrentlyInMaintenance As Boolean
End Type


'------------------------------------------------------------------------------
' Purpose : Calculates the maintenance-window status for one aircraft.
' Input : settings - validated planner settings.
' weekStartDates - visible planner week start dates.
' weekCyclePosition - weekly cycle position array for this aircraft.
' currentWeekIndex - current visible week index, or 0 if not visible.
' isDownMaintenance - True when aircraft is in long-term maintenance.
' downStartValue - Rotation down/maintenance start value.
' expectedRtsValue - Rotation expected RTS / ETBOL value.
' rtsFlyStart - first Monday after ETBOL.
' Output : MaintenanceWindowStatus.
'------------------------------------------------------------------------------
Public Function GetMaintenanceWindowStatus(ByRef settings As PlannerSettings, _
                                           ByRef weekStartDates() As Date, _
                                           ByRef weekCyclePosition() As Long, _
                                           ByVal currentWeekIndex As Long, _
                                           ByVal isDownMaintenance As Boolean, _
                                           ByVal downStartValue As Variant, _
                                           ByVal expectedRtsValue As Variant, _
                                           ByVal rtsFlyStart As Date) As MaintenanceWindowStatus

    Dim status As MaintenanceWindowStatus

    If isDownMaintenance Then

        status = GetLongTermDownMaintenanceStatus( _
            settings, _
            downStartValue, _
            expectedRtsValue, _
            rtsFlyStart)

    Else

        status = GetInCycleMaintenanceStatus( _
            settings, _
            weekStartDates, _
            weekCyclePosition, _
            currentWeekIndex)

    End If

    GetMaintenanceWindowStatus = status

End Function


'------------------------------------------------------------------------------
' Purpose : Calculates maintenance status for a long-term down aircraft.
'------------------------------------------------------------------------------
Private Function GetLongTermDownMaintenanceStatus(ByRef settings As PlannerSettings, _
                                                  ByVal downStartValue As Variant, _
                                                  ByVal expectedRtsValue As Variant, _
                                                  ByVal rtsFlyStart As Date) As MaintenanceWindowStatus

    Dim status As MaintenanceWindowStatus

    Dim downStartDate As Date
    Dim expectedRtsDate As Date

    downStartDate = GetDateOrDefault(downStartValue, Date)
    expectedRtsDate = GetDateOrDefault(expectedRtsValue, Date + 56)

    If expectedRtsDate < downStartDate Then
        expectedRtsDate = downStartDate + 56
    End If

    If expectedRtsDate >= Date Then
        status.daysToMaintenance = "ETBOL: " & Format$(expectedRtsDate, "DD/MM")
    Else
        status.daysToMaintenance = "ETBOL overdue"
    End If

    status.IsCurrentlyInMaintenance = (Date < rtsFlyStart)

    ' Long-term down aircraft use actual ETBOL dates rather than cycle week
    ' indexes for the dashboard summary. The start/end indexes are left as zero.
    status.startWeekIndex = 0
    status.endWeekIndex = 0

    GetLongTermDownMaintenanceStatus = status

End Function


'------------------------------------------------------------------------------
' Purpose : Calculates maintenance status for a normal in-cycle aircraft.
'------------------------------------------------------------------------------
Private Function GetInCycleMaintenanceStatus(ByRef settings As PlannerSettings, _
                                             ByRef weekStartDates() As Date, _
                                             ByRef weekCyclePosition() As Long, _
                                             ByVal currentWeekIndex As Long) As MaintenanceWindowStatus

    Dim status As MaintenanceWindowStatus

    If Not IsDisplayWeekIndexValid(currentWeekIndex, settings.displayWeeksShown) Then
        status.daysToMaintenance = "-"
        status.IsCurrentlyInMaintenance = False
        GetInCycleMaintenanceStatus = status
        Exit Function
    End If

    If weekCyclePosition(currentWeekIndex) > settings.flyingWeeks Then

        status.startWeekIndex = FindCurrentMaintenanceStart( _
            weekCyclePosition, _
            currentWeekIndex, _
            settings.flyingWeeks)

        status.endWeekIndex = FindCurrentMaintenanceEnd( _
            weekCyclePosition, _
            currentWeekIndex, _
            settings.displayWeeksShown, _
            settings.flyingWeeks)

        status.daysToMaintenance = 0
        status.IsCurrentlyInMaintenance = True

    Else

        status.startWeekIndex = FindNextMaintenanceStart( _
            weekCyclePosition, _
            currentWeekIndex, _
            settings.displayWeeksShown, _
            settings.flyingWeeks)

        If status.startWeekIndex > 0 Then
            status.endWeekIndex = FindCurrentMaintenanceEnd( _
                weekCyclePosition, _
                status.startWeekIndex, _
                settings.displayWeeksShown, _
                settings.flyingWeeks)

            status.daysToMaintenance = CLng(weekStartDates(status.startWeekIndex) - Date)

            If status.daysToMaintenance < 0 Then
                status.daysToMaintenance = 0
            End If
        Else
            status.daysToMaintenance = "-"
        End If

        status.IsCurrentlyInMaintenance = False

    End If

    GetInCycleMaintenanceStatus = status

End Function


'------------------------------------------------------------------------------
' Purpose : Finds the first week index of the current maintenance block.
'------------------------------------------------------------------------------
Private Function FindCurrentMaintenanceStart(ByRef weekCyclePosition() As Long, _
                                             ByVal currentWeekIndex As Long, _
                                             ByVal flyingWeeks As Long) As Long

    Dim weekIndex As Long

    For weekIndex = currentWeekIndex To 1 Step -1

        If weekCyclePosition(weekIndex) > flyingWeeks Then
            FindCurrentMaintenanceStart = weekIndex
        Else
            Exit For
        End If

    Next weekIndex

End Function


'------------------------------------------------------------------------------
' Purpose : Finds the last week index of the maintenance block containing start.
'------------------------------------------------------------------------------
Private Function FindCurrentMaintenanceEnd(ByRef weekCyclePosition() As Long, _
                                           ByVal startWeekIndex As Long, _
                                           ByVal displayWeeksShown As Long, _
                                           ByVal flyingWeeks As Long) As Long

    Dim weekIndex As Long

    For weekIndex = startWeekIndex To displayWeeksShown

        If weekCyclePosition(weekIndex) > flyingWeeks Then
            FindCurrentMaintenanceEnd = weekIndex
        Else
            Exit For
        End If

    Next weekIndex

End Function


'------------------------------------------------------------------------------
' Purpose : Finds the first week index of the next maintenance block.
'------------------------------------------------------------------------------
Private Function FindNextMaintenanceStart(ByRef weekCyclePosition() As Long, _
                                          ByVal currentWeekIndex As Long, _
                                          ByVal displayWeeksShown As Long, _
                                          ByVal flyingWeeks As Long) As Long

    Dim weekIndex As Long

    For weekIndex = currentWeekIndex To displayWeeksShown

        If weekCyclePosition(weekIndex) > flyingWeeks Then
            FindNextMaintenanceStart = weekIndex
            Exit Function
        End If

    Next weekIndex

    FindNextMaintenanceStart = 0

End Function


'------------------------------------------------------------------------------
' Purpose : Returns a valid date or a fallback date.
'------------------------------------------------------------------------------
Private Function GetDateOrDefault(ByVal rawValue As Variant, _
                                  ByVal fallbackDate As Date) As Date

    If IsDate(rawValue) Then
        GetDateOrDefault = CDate(rawValue)
    Else
        GetDateOrDefault = fallbackDate
    End If

End Function

