Attribute VB_Name = "modAircraftRoster"
Option Explicit

'------------------------------------------------------------------------------
' Module : modAircraftRoster
' Purpose : Loads aircraft planning information from the Rotation sheet.
'------------------------------------------------------------------------------

Private Const ROTATION_SHEET_NAME As String = "Rotation"

Private Const ROTATION_FIRST_SLOT_ROW As Long = 5

' Column headers
Private Const ROTATION_COL_SLOT As Long = 1
Private Const ROTATION_COL_TAIL_NUMBER As Long = 2
Private Const ROTATION_COL_PLANNING_MODE As Long = 3
Private Const ROTATION_COL_CYCLE_REFERENCE_DATE As Long = 4
Private Const ROTATION_COL_CYCLE_WEEK_AT_REFERENCE As Long = 5
Private Const ROTATION_COL_CYCLE_WEEK As Long = 6
Private Const ROTATION_COL_DOWN_START As Long = 7
Private Const ROTATION_COL_ETBOL As Long = 8
Private Const ROTATION_COL_MAINTENANCE_REASON As Long = 9
Private Const ROTATION_COL_CYCLE_WEEK_ON_RETURN As Long = 10
Private Const ROTATION_COL_CYCLE_NOTES As Long = 11

Public Const PLANNING_MODE_CYCLE As String = "CYCLE"
Public Const PLANNING_MODE_DOWN As String = "DOWN"

'------------------------------------------------------------------------------
' Type : AircraftRoster
' Purpose : Stores aircraft planning data loaded from the Rotation sheet.
' Notes : Arrays are used for compatibility with the existing RefreshAll code.
'------------------------------------------------------------------------------
Public Type aircraftRoster
    aircraftCount As Long

    tailNumbers() As String
    cycleStarts() As Long
    planningModes() As String
    MaintenanceStarts() As Variant
    ExpectedReturnToServiceDates() As Variant
    MaintenanceReasons() As String
End Type


'------------------------------------------------------------------------------
' Purpose : Loads all aircraft from the Rotation sheet.
' Input : settings - validated planner settings.
' Output : AircraftRoster containing all populated aircraft slots.
' Raises : Error if a Down Maintenance aircraft has invalid cycle setup.
'------------------------------------------------------------------------------
Public Function LoadAircraftRoster(ByRef settings As PlannerSettings) As aircraftRoster

    Dim rotationWs As Worksheet
    Set rotationWs = ThisWorkbook.Worksheets(ROTATION_SHEET_NAME)

    Dim roster As aircraftRoster

    ReDim roster.tailNumbers(1 To settings.MaxAircraftSlots)
    ReDim roster.cycleStarts(1 To settings.MaxAircraftSlots)
    ReDim roster.planningModes(1 To settings.MaxAircraftSlots)
    ReDim roster.MaintenanceStarts(1 To settings.MaxAircraftSlots)
    ReDim roster.ExpectedReturnToServiceDates(1 To settings.MaxAircraftSlots)
    ReDim roster.MaintenanceReasons(1 To settings.MaxAircraftSlots)

    Dim slotIndex As Long

    For slotIndex = 1 To settings.MaxAircraftSlots

        Dim rotationRow As Long
        rotationRow = ROTATION_FIRST_SLOT_ROW + slotIndex - 1

        Dim tailNumber As String
        tailNumber = Trim$(CStr(rotationWs.Cells(rotationRow, ROTATION_COL_TAIL_NUMBER).Value))

        If Len(tailNumber) > 0 Then
            AddAircraftToRoster roster, settings, rotationWs, rotationRow, tailNumber
        End If

    Next slotIndex

    SortAircraftRoster roster, settings.DashboardSortOrder

    LoadAircraftRoster = roster

End Function


'------------------------------------------------------------------------------
' Purpose : Adds one aircraft row from Rotation into the roster.
' Input : roster - roster being populated.
' settings - validated planner settings.
' rotationWs - Rotation worksheet.
' rowNumber - worksheet row containing the aircraft.
' tailNumber - aircraft tail number.
'------------------------------------------------------------------------------
Private Sub AddAircraftToRoster(ByRef roster As aircraftRoster, _
                                ByRef settings As PlannerSettings, _
                                ByVal rotationWs As Worksheet, _
                                ByVal rowNumber As Long, _
                                ByVal tailNumber As String)

    Dim aircraftIndex As Long
    Dim cycleWeekSourceValue As Variant
    
    roster.aircraftCount = roster.aircraftCount + 1
    
    aircraftIndex = roster.aircraftCount

    roster.tailNumbers(aircraftIndex) = tailNumber
    roster.planningModes(aircraftIndex) = NormalisePlanningMode( _
        CStr(rotationWs.Cells(rowNumber, ROTATION_COL_PLANNING_MODE).Value))
    
    If roster.planningModes(aircraftIndex) = PLANNING_MODE_DOWN Then
        cycleWeekSourceValue = rotationWs.Cells(rowNumber, ROTATION_COL_CYCLE_WEEK_ON_RETURN).Value
    Else
        cycleWeekSourceValue = rotationWs.Cells(rowNumber, ROTATION_COL_CYCLE_WEEK).Value
    End If
    
    roster.cycleStarts(aircraftIndex) = GetValidCycleWeek( _
        tailNumber, _
        roster.planningModes(aircraftIndex), _
        cycleWeekSourceValue, _
        settings.cycleLength)

    roster.MaintenanceStarts(aircraftIndex) = _
        rotationWs.Cells(rowNumber, ROTATION_COL_DOWN_START).Value

    roster.ExpectedReturnToServiceDates(aircraftIndex) = _
        rotationWs.Cells(rowNumber, ROTATION_COL_ETBOL).Value

    roster.MaintenanceReasons(aircraftIndex) = _
        Trim$(CStr(rotationWs.Cells(rowNumber, ROTATION_COL_MAINTENANCE_REASON).Value))

End Sub


'------------------------------------------------------------------------------
' Purpose : Converts user-facing Rotation mode text into internal mode codes.
' Input : modeText - value from Rotation column C.
' Output : PLANNING_MODE_CYCLE or PLANNING_MODE_DOWN.
'------------------------------------------------------------------------------
Private Function NormalisePlanningMode(ByVal modeText As String) As String

    Select Case UCase$(Trim$(modeText))

        Case "IN CYCLE", "CYCLE", "Y", "YES", "ACTIVE"
            NormalisePlanningMode = PLANNING_MODE_CYCLE

        Case "DOWN MAINTENANCE", "DOWN", "D", "LONG DOWN", "LONG_DOWN"
            NormalisePlanningMode = PLANNING_MODE_DOWN

        Case Else
            NormalisePlanningMode = PLANNING_MODE_CYCLE

    End Select

End Function


'------------------------------------------------------------------------------
' Purpose : Validates and returns the aircraft cycle week.
' Input : tailNumber - aircraft tail number, used in error messages.
' planningMode - internal planning mode.
' rawValue - Rotation cycle week value.
' cycleLength - total flying + maintenance cycle length.
' Output : Valid cycle week.
' Raises : Error if Down Maintenance aircraft has blank/invalid cycle week.
'------------------------------------------------------------------------------
Private Function GetValidCycleWeek(ByVal tailNumber As String, _
                                   ByVal planningMode As String, _
                                   ByVal rawValue As Variant, _
                                   ByVal cycleLength As Long) As Long

    If planningMode = PLANNING_MODE_DOWN Then

        If Not IsNumeric(rawValue) Or IsEmpty(rawValue) Then
            Err.Raise vbObjectError + 2300, _
                      "GetValidCycleWeek", _
                      "Rotation setup issue for " & tailNumber & "." & vbCrLf & vbCrLf & _
                      "This aircraft is marked as Down Maintenance, but its Cycle Week on Return is blank." & vbCrLf & vbCrLf & _
                      "For down aircraft, Cycle Week on Return means the cycle position the aircraft will slot into after the ETBOL." & vbCrLf & _
                      "Enter a value from 1 to " & cycleLength & "."
        End If

        GetValidCycleWeek = CLng(rawValue)

        If GetValidCycleWeek < 1 Or GetValidCycleWeek > cycleLength Then
            Err.Raise vbObjectError + 2301, _
                      "GetValidCycleWeek", _
                      "Rotation setup issue for " & tailNumber & "." & vbCrLf & vbCrLf & _
                      "This aircraft is marked as Down Maintenance, but its Cycle Week on Return is outside the valid range." & vbCrLf & _
                      "For down aircraft, Cycle Week on Return means the cycle position the aircraft will slot into after the ETBOL." & vbCrLf & vbCrLf & _
                      "Enter a value from 1 to " & cycleLength & "."
        End If

    Else

        If IsNumeric(rawValue) And Not IsEmpty(rawValue) Then
            GetValidCycleWeek = CLng(rawValue)

            If GetValidCycleWeek < 1 Or GetValidCycleWeek > cycleLength Then
                GetValidCycleWeek = 1
            End If
        Else
            GetValidCycleWeek = 1
        End If

    End If

End Function


'------------------------------------------------------------------------------
' Purpose : Returns the Rotation sheet row for a roster slot index.
' Input : slotIndex - roster slot number (1-based).
' Output : Worksheet row number on the Rotation sheet.
'------------------------------------------------------------------------------
Public Function GetRotationRowForSlot(ByVal slotIndex As Long) As Long

    GetRotationRowForSlot = ROTATION_FIRST_SLOT_ROW + slotIndex - 1

End Function


'------------------------------------------------------------------------------
' Purpose : Writes a new in-cycle aircraft row on the Rotation sheet.
' Input : slotIndex - roster slot number (1-based).
' tailNumber - aircraft tail number.
' cycleWeekAtReference - cycle week at the reference date.
' Notes : Column F (calculated cycle week) is left alone so sheet formulas can recalculate.
'------------------------------------------------------------------------------
Public Sub WriteNewAircraftRotationRow(ByVal slotIndex As Long, _
                                       ByVal tailNumber As String, _
                                       ByVal cycleWeekAtReference As Long)

    Dim rotationWs As Worksheet
    Set rotationWs = ThisWorkbook.Worksheets(ROTATION_SHEET_NAME)

    Dim rotationRow As Long
    rotationRow = ROTATION_FIRST_SLOT_ROW + slotIndex - 1

    rotationWs.Cells(rotationRow, ROTATION_COL_SLOT).Value = slotIndex
    rotationWs.Cells(rotationRow, ROTATION_COL_TAIL_NUMBER).Value = UCase$(Trim$(tailNumber))
    rotationWs.Cells(rotationRow, ROTATION_COL_PLANNING_MODE).Value = "In Cycle"
    rotationWs.Cells(rotationRow, ROTATION_COL_PLANNING_MODE).HorizontalAlignment = xlCenter

    rotationWs.Cells(rotationRow, ROTATION_COL_CYCLE_REFERENCE_DATE).Value = GetNamedDate("ForecastStart")
    rotationWs.Cells(rotationRow, ROTATION_COL_CYCLE_REFERENCE_DATE).NumberFormat = "dd mmm yyyy"

    rotationWs.Cells(rotationRow, ROTATION_COL_CYCLE_WEEK_AT_REFERENCE).Value = cycleWeekAtReference
    rotationWs.Cells(rotationRow, ROTATION_COL_CYCLE_WEEK_AT_REFERENCE).HorizontalAlignment = xlCenter

    rotationWs.Cells(rotationRow, ROTATION_COL_DOWN_START).ClearContents
    rotationWs.Cells(rotationRow, ROTATION_COL_ETBOL).ClearContents
    rotationWs.Cells(rotationRow, ROTATION_COL_MAINTENANCE_REASON).ClearContents
    rotationWs.Cells(rotationRow, ROTATION_COL_CYCLE_WEEK_ON_RETURN).ClearContents
    rotationWs.Cells(rotationRow, ROTATION_COL_CYCLE_NOTES).ClearContents

    rotationWs.Calculate

End Sub


'------------------------------------------------------------------------------
' Purpose : Clears aircraft data from a Rotation row while preserving the slot label.
' Input : slotIndex - roster slot number (1-based).
' Notes : Column A (slot) and column F (calculated cycle week formula) are preserved.
'------------------------------------------------------------------------------
Public Sub ClearRotationRowForRemovedAircraft(ByVal slotIndex As Long)

    Dim rotationWs As Worksheet
    Set rotationWs = ThisWorkbook.Worksheets(ROTATION_SHEET_NAME)

    Dim rotationRow As Long
    rotationRow = ROTATION_FIRST_SLOT_ROW + slotIndex - 1

    rotationWs.Range( _
        rotationWs.Cells(rotationRow, ROTATION_COL_TAIL_NUMBER), _
        rotationWs.Cells(rotationRow, ROTATION_COL_CYCLE_WEEK_AT_REFERENCE)).ClearContents

    rotationWs.Range( _
        rotationWs.Cells(rotationRow, ROTATION_COL_DOWN_START), _
        rotationWs.Cells(rotationRow, ROTATION_COL_CYCLE_NOTES)).ClearContents

    rotationWs.Cells(rotationRow, ROTATION_COL_SLOT).Value = slotIndex

    rotationWs.Calculate

End Sub


'------------------------------------------------------------------------------
' Purpose : Closes gaps on the Rotation sheet by moving occupied rows upward.
' Input : maxSlots - maximum aircraft slots configured for the planner.
' Notes : Column F formulas remain on each row; only user-editable values move.
'------------------------------------------------------------------------------
Public Sub CompactRotationSlots(ByVal maxSlots As Long)

    Dim rotationWs As Worksheet
    Set rotationWs = ThisWorkbook.Worksheets(ROTATION_SHEET_NAME)

    Dim filledCount As Long
    Dim slotIndex As Long

    For slotIndex = 1 To maxSlots

        If RotationSlotHasTail(rotationWs, slotIndex) Then
            filledCount = filledCount + 1

            If slotIndex <> filledCount Then
                MoveRotationSlotData rotationWs, slotIndex, filledCount
            End If
        End If

    Next slotIndex

    For slotIndex = filledCount + 1 To maxSlots
        ClearRotationRowForRemovedAircraft slotIndex
    Next slotIndex

    rotationWs.Calculate

End Sub


'------------------------------------------------------------------------------
' Purpose : Returns the next Rotation slot for a newly added aircraft.
' Input : maxSlots - maximum aircraft slots configured for the planner.
' Output : Slot index after the last occupied slot, or 0 if the roster is full.
'------------------------------------------------------------------------------
Public Function GetNextRotationSlotForAdd(ByVal maxSlots As Long) As Long

    CompactRotationSlots maxSlots

    Dim rotationWs As Worksheet
    Set rotationWs = ThisWorkbook.Worksheets(ROTATION_SHEET_NAME)

    Dim slotIndex As Long
    Dim filledCount As Long

    For slotIndex = 1 To maxSlots
        If RotationSlotHasTail(rotationWs, slotIndex) Then
            filledCount = filledCount + 1
        End If
    Next slotIndex

    If filledCount >= maxSlots Then
        GetNextRotationSlotForAdd = 0
    Else
        GetNextRotationSlotForAdd = filledCount + 1
    End If

End Function


'------------------------------------------------------------------------------
' Purpose : Sorts the loaded roster according to the dashboard order setting.
' Input : roster - roster loaded from the Rotation sheet.
' sortOrder - DASHBOARD_SORT_TAIL or DASHBOARD_SORT_CYCLE.
'------------------------------------------------------------------------------
Private Sub SortAircraftRoster(ByRef roster As aircraftRoster, ByVal sortOrder As String)

    If roster.aircraftCount <= 1 Then Exit Sub

    Dim firstIndex As Long
    Dim secondIndex As Long

    For firstIndex = 1 To roster.aircraftCount - 1

        For secondIndex = firstIndex + 1 To roster.aircraftCount

            Dim shouldSwap As Boolean

            If StrComp(sortOrder, DASHBOARD_SORT_CYCLE, vbTextCompare) = 0 Then
                shouldSwap = (CompareRosterByCycleWeek(roster, firstIndex, secondIndex) > 0)
            Else
                shouldSwap = (StrComp(UCase$(roster.tailNumbers(firstIndex)), _
                                       UCase$(roster.tailNumbers(secondIndex)), _
                                       vbTextCompare) > 0)
            End If

            If shouldSwap Then
                SwapRosterEntries roster, firstIndex, secondIndex
            End If

        Next secondIndex

    Next firstIndex

End Sub


'------------------------------------------------------------------------------
' Purpose : Compares two roster entries for cycle-week dashboard ordering.
' Notes : In-cycle aircraft sort by cycle week; down aircraft sort last by ETBOL.
'------------------------------------------------------------------------------
Private Function CompareRosterByCycleWeek(ByRef roster As aircraftRoster, _
                                          ByVal firstIndex As Long, _
                                          ByVal secondIndex As Long) As Long

    Dim firstIsDown As Boolean
    Dim secondIsDown As Boolean

    firstIsDown = (roster.planningModes(firstIndex) = PLANNING_MODE_DOWN)
    secondIsDown = (roster.planningModes(secondIndex) = PLANNING_MODE_DOWN)

    If firstIsDown And Not secondIsDown Then
        CompareRosterByCycleWeek = 1
        Exit Function
    End If

    If Not firstIsDown And secondIsDown Then
        CompareRosterByCycleWeek = -1
        Exit Function
    End If

    If firstIsDown And secondIsDown Then
        CompareRosterByCycleWeek = CompareSortDates( _
            roster.ExpectedReturnToServiceDates(firstIndex), _
            roster.ExpectedReturnToServiceDates(secondIndex))
        Exit Function
    End If

    If roster.cycleStarts(firstIndex) < roster.cycleStarts(secondIndex) Then
        CompareRosterByCycleWeek = -1
    ElseIf roster.cycleStarts(firstIndex) > roster.cycleStarts(secondIndex) Then
        CompareRosterByCycleWeek = 1
    Else
        CompareRosterByCycleWeek = StrComp(UCase$(roster.tailNumbers(firstIndex)), _
                                            UCase$(roster.tailNumbers(secondIndex)), _
                                            vbTextCompare)
    End If

End Function


'------------------------------------------------------------------------------
' Purpose : Compares two optional dates for roster sorting.
' Notes : Blank dates sort after valid dates.
'------------------------------------------------------------------------------
Private Function CompareSortDates(ByVal firstDate As Variant, ByVal secondDate As Variant) As Long

    Dim firstHasDate As Boolean
    Dim secondHasDate As Boolean

    firstHasDate = IsDate(firstDate) And Not IsEmpty(firstDate)
    secondHasDate = IsDate(secondDate) And Not IsEmpty(secondDate)

    If Not firstHasDate And Not secondHasDate Then
        CompareSortDates = 0
        Exit Function
    End If

    If Not firstHasDate Then
        CompareSortDates = 1
        Exit Function
    End If

    If Not secondHasDate Then
        CompareSortDates = -1
        Exit Function
    End If

    If CDate(firstDate) < CDate(secondDate) Then
        CompareSortDates = -1
    ElseIf CDate(firstDate) > CDate(secondDate) Then
        CompareSortDates = 1
    Else
        CompareSortDates = 0
    End If

End Function


'------------------------------------------------------------------------------
' Purpose : Swaps two roster entries in all parallel arrays.
'------------------------------------------------------------------------------
Private Sub SwapRosterEntries(ByRef roster As aircraftRoster, _
                              ByVal firstIndex As Long, _
                              ByVal secondIndex As Long)

    Dim tempTail As String
    Dim tempCycleStart As Long
    Dim tempPlanningMode As String
    Dim tempMaintenanceStart As Variant
    Dim tempExpectedRts As Variant
    Dim tempMaintenanceReason As String

    tempTail = roster.tailNumbers(firstIndex)
    tempCycleStart = roster.cycleStarts(firstIndex)
    tempPlanningMode = roster.planningModes(firstIndex)
    tempMaintenanceStart = roster.MaintenanceStarts(firstIndex)
    tempExpectedRts = roster.ExpectedReturnToServiceDates(firstIndex)
    tempMaintenanceReason = roster.MaintenanceReasons(firstIndex)

    roster.tailNumbers(firstIndex) = roster.tailNumbers(secondIndex)
    roster.cycleStarts(firstIndex) = roster.cycleStarts(secondIndex)
    roster.planningModes(firstIndex) = roster.planningModes(secondIndex)
    roster.MaintenanceStarts(firstIndex) = roster.MaintenanceStarts(secondIndex)
    roster.ExpectedReturnToServiceDates(firstIndex) = roster.ExpectedReturnToServiceDates(secondIndex)
    roster.MaintenanceReasons(firstIndex) = roster.MaintenanceReasons(secondIndex)

    roster.tailNumbers(secondIndex) = tempTail
    roster.cycleStarts(secondIndex) = tempCycleStart
    roster.planningModes(secondIndex) = tempPlanningMode
    roster.MaintenanceStarts(secondIndex) = tempMaintenanceStart
    roster.ExpectedReturnToServiceDates(secondIndex) = tempExpectedRts
    roster.MaintenanceReasons(secondIndex) = tempMaintenanceReason

End Sub


'------------------------------------------------------------------------------
' Purpose : Returns True when a Rotation slot contains a tail number.
'------------------------------------------------------------------------------
Private Function RotationSlotHasTail(ByVal rotationWs As Worksheet, ByVal slotIndex As Long) As Boolean

    RotationSlotHasTail = _
        Len(Trim$(CStr(rotationWs.Cells(GetRotationRowForSlot(slotIndex), ROTATION_COL_TAIL_NUMBER).Value))) > 0

End Function


'------------------------------------------------------------------------------
' Purpose : Copies user-editable Rotation data from one slot row to another.
' Input : sourceSlot - slot to copy from.
' targetSlot - slot to copy to.
' Notes : Column F formulas stay on each row and recalculate after the move.
'------------------------------------------------------------------------------
Private Sub MoveRotationSlotData(ByVal rotationWs As Worksheet, _
                                 ByVal sourceSlot As Long, _
                                 ByVal targetSlot As Long)

    Dim sourceRow As Long
    Dim targetRow As Long

    sourceRow = GetRotationRowForSlot(sourceSlot)
    targetRow = GetRotationRowForSlot(targetSlot)

    rotationWs.Cells(targetRow, ROTATION_COL_SLOT).Value = targetSlot
    rotationWs.Cells(targetRow, ROTATION_COL_TAIL_NUMBER).Value = _
        rotationWs.Cells(sourceRow, ROTATION_COL_TAIL_NUMBER).Value
    rotationWs.Cells(targetRow, ROTATION_COL_PLANNING_MODE).Value = _
        rotationWs.Cells(sourceRow, ROTATION_COL_PLANNING_MODE).Value
    rotationWs.Cells(targetRow, ROTATION_COL_CYCLE_REFERENCE_DATE).Value = _
        rotationWs.Cells(sourceRow, ROTATION_COL_CYCLE_REFERENCE_DATE).Value
    rotationWs.Cells(targetRow, ROTATION_COL_CYCLE_WEEK_AT_REFERENCE).Value = _
        rotationWs.Cells(sourceRow, ROTATION_COL_CYCLE_WEEK_AT_REFERENCE).Value
    rotationWs.Cells(targetRow, ROTATION_COL_DOWN_START).Value = _
        rotationWs.Cells(sourceRow, ROTATION_COL_DOWN_START).Value
    rotationWs.Cells(targetRow, ROTATION_COL_ETBOL).Value = _
        rotationWs.Cells(sourceRow, ROTATION_COL_ETBOL).Value
    rotationWs.Cells(targetRow, ROTATION_COL_MAINTENANCE_REASON).Value = _
        rotationWs.Cells(sourceRow, ROTATION_COL_MAINTENANCE_REASON).Value
    rotationWs.Cells(targetRow, ROTATION_COL_CYCLE_WEEK_ON_RETURN).Value = _
        rotationWs.Cells(sourceRow, ROTATION_COL_CYCLE_WEEK_ON_RETURN).Value
    rotationWs.Cells(targetRow, ROTATION_COL_CYCLE_NOTES).Value = _
        rotationWs.Cells(sourceRow, ROTATION_COL_CYCLE_NOTES).Value

    ClearRotationRowForRemovedAircraft sourceSlot

End Sub

