Attribute VB_Name = "modChartCellAggregation"
Option Explicit

'------------------------------------------------------------------------------
' Module : modChartCellAggregation
' Purpose : Aggregates task occurrences into chart cells before writing them
' to the aircraft chart worksheet.
'------------------------------------------------------------------------------

Public Type ChartCellAggregation
    CellCounts As Object
    CellTasks As Object
    CellGrouped As Object
    CellPushed As Object
    CellPulled As Object
    CellHighlighted As Object
End Type


'------------------------------------------------------------------------------
' Purpose : Creates an empty chart cell aggregation result.
'------------------------------------------------------------------------------
Public Function CreateChartCellAggregation() As ChartCellAggregation

    Dim aggregation As ChartCellAggregation

    Set aggregation.CellCounts = CreateObject("Scripting.Dictionary")
    Set aggregation.CellTasks = CreateObject("Scripting.Dictionary")
    Set aggregation.CellGrouped = CreateObject("Scripting.Dictionary")
    Set aggregation.CellPushed = CreateObject("Scripting.Dictionary")
    Set aggregation.CellPulled = CreateObject("Scripting.Dictionary")
    Set aggregation.CellHighlighted = CreateObject("Scripting.Dictionary")

    CreateChartCellAggregation = aggregation

End Function


'------------------------------------------------------------------------------
' Purpose : Aggregates task occurrences into chart cell dictionaries.
' Input : taskOccurrences - populated task occurrence records.
' totalOccurrences - number of populated task occurrences.
' rowLookup - dictionary mapping interval pair key to row.
' colLookup - dictionary mapping week start date to column.
' settings - planner settings.
' displayStart - first visible display week.
' weekCyclePositions - aircraft weekly cycle positions.
' Output : ChartCellAggregation with cell counts, comment text, and flags.
'------------------------------------------------------------------------------
Public Function BuildChartCellAggregation(ByRef taskOccurrences() As TaskOccurrence, _
                                          ByVal totalOccurrences As Long, _
                                          ByVal rowLookup As Object, _
                                          ByVal colLookup As Object, _
                                          ByRef settings As PlannerSettings, _
                                          ByVal displayStart As Date, _
                                          ByRef weekCyclePositions() As Long) As ChartCellAggregation

    Dim aggregation As ChartCellAggregation
    aggregation = CreateChartCellAggregation()

    Dim chartTaskCodes As Object
    Set chartTaskCodes = CreateObject("Scripting.Dictionary")

    Dim occurrenceIndex As Long

    For occurrenceIndex = 1 To totalOccurrences

        Dim pairKey As String
        pairKey = BuildChartPairKey( _
            taskOccurrences(occurrenceIndex).intervalType, _
            taskOccurrences(occurrenceIndex).intervalValue)

        If Not rowLookup.Exists(pairKey) Then GoTo NextOccurrence

        Dim chartRow As Long
        chartRow = CLng(rowLookup(pairKey))

        Dim weekKey As Long
        weekKey = CLng(taskOccurrences(occurrenceIndex).scheduledWeek)

        If Not colLookup.Exists(weekKey) Then GoTo NextOccurrence

        Dim chartColumn As Long
        chartColumn = CLng(colLookup(weekKey))

        Dim cellKey As String
        cellKey = BuildChartCellKey(chartRow, chartColumn)

        Dim chartPeriodKey As String
        chartPeriodKey = BuildChartPeriodKey( _
            chartRow, _
            chartColumn, _
            taskOccurrences(occurrenceIndex).scheduledWeek, _
            displayStart, _
            settings, _
            weekCyclePositions)

        Dim chartTaskKey As String
        chartTaskKey = BuildChartTaskDeduplicationKey( _
            chartPeriodKey, _
            taskOccurrences(occurrenceIndex))

        If chartTaskCodes.Exists(chartTaskKey) Then GoTo NextOccurrence
        chartTaskCodes.Add chartTaskKey, True

        AddOccurrenceToChartAggregation _
            aggregation, _
            cellKey, _
            taskOccurrences(occurrenceIndex)

NextOccurrence:
    Next occurrenceIndex

    BuildChartCellAggregation = aggregation

End Function


'------------------------------------------------------------------------------
' Purpose : Adds one occurrence into the aggregation dictionaries.
'------------------------------------------------------------------------------
Private Sub AddOccurrenceToChartAggregation(ByRef aggregation As ChartCellAggregation, _
                                            ByVal cellKey As String, _
                                            ByRef occurrence As TaskOccurrence)

    Dim taskLine As String
    taskLine = BuildChartOccurrenceCommentLine(occurrence)

    If aggregation.CellCounts.Exists(cellKey) Then

        aggregation.CellCounts(cellKey) = aggregation.CellCounts(cellKey) + 1

        aggregation.CellTasks(cellKey) = _
            aggregation.CellTasks(cellKey) & vbLf & vbLf & _
            CStr(aggregation.CellCounts(cellKey)) & ". " & taskLine

        If occurrence.GroupedDirection <> 0 Then
            aggregation.CellGrouped(cellKey) = occurrence.GroupedDirection
        End If

        If occurrence.wasAutoPushed Then aggregation.CellPushed(cellKey) = True
        If occurrence.wasAutoPulled Then aggregation.CellPulled(cellKey) = True
        If occurrence.isHighlighted Then aggregation.CellHighlighted(cellKey) = True

    Else

        aggregation.CellCounts.Add cellKey, 1
        aggregation.CellTasks.Add cellKey, "1. " & taskLine
        aggregation.CellGrouped.Add cellKey, occurrence.GroupedDirection
        aggregation.CellPushed.Add cellKey, occurrence.wasAutoPushed
        aggregation.CellPulled.Add cellKey, occurrence.wasAutoPulled
        aggregation.CellHighlighted.Add cellKey, occurrence.isHighlighted

    End If

End Sub


'------------------------------------------------------------------------------
' Purpose : Builds one chart comment line for an occurrence.
'------------------------------------------------------------------------------
Public Function BuildChartOccurrenceCommentLine(ByRef occurrence As TaskOccurrence) As String

    Dim shortDescription As String
    shortDescription = TruncateDesc(occurrence.taskDescription)

    Dim taskLine As String

    If occurrence.isHighlighted Then
        taskLine = "*** " & occurrence.taskCode & " - " & shortDescription & " ***"
    Else
        taskLine = occurrence.taskCode & " - " & shortDescription
    End If

    taskLine = taskLine & RightAlignTag("[" & BuildChartDueTimeText(occurrence) & "]")

    If occurrence.GroupedDirection = -1 Then
        taskLine = taskLine & RightAlignTag( _
            "[PULLED TO ALIGN by " & FormatOccurrenceGroupedAlignmentAmount(occurrence) & "]")
    ElseIf occurrence.GroupedDirection = 1 Then
        taskLine = taskLine & RightAlignTag( _
            "[PUSHED TO ALIGN by " & FormatOccurrenceGroupedAlignmentAmount(occurrence) & "]")
    End If

    If occurrence.wasAutoPulled Then
        taskLine = taskLine & RightAlignTag( _
            "[PULLED " & FormatOccurrencePullAmount(occurrence) & _
            " (" & Format$(occurrence.pullPercentUsed * 100, "0") & _
            "%) to Maint]")
    End If

    If occurrence.wasAutoPushed Then
        taskLine = taskLine & RightAlignTag(BuildChartExtensionTag(occurrence))
    End If

    BuildChartOccurrenceCommentLine = taskLine

End Function


'------------------------------------------------------------------------------
' Purpose : Builds the due-time text shown in chart comments.
'------------------------------------------------------------------------------
Private Function BuildChartDueTimeText(ByRef occurrence As TaskOccurrence) As String

    If occurrence.intervalType = "DD" Then

        BuildChartDueTimeText = "Due " & _
                                Format$(occurrence.firstDueDate, "DD/MM/YYYY")

    Else

        If occurrence.lifeRemaining < 0 Then
            BuildChartDueTimeText = "Overdue by " & _
                                    FormatHoursToHMM(Abs(occurrence.lifeRemaining)) & _
                                    occurrence.intervalType
        Else
            BuildChartDueTimeText = "Due in " & _
                                    FormatHoursToHMM(occurrence.lifeRemaining) & _
                                    occurrence.intervalType
        End If

    End If

End Function


'------------------------------------------------------------------------------
' Purpose : Builds the extension-required tag shown in chart comments.
'------------------------------------------------------------------------------
Public Function BuildChartExtensionTag(ByRef occurrence As TaskOccurrence) As String

    If occurrence.existingExtensionPercent > 0 Then

        If occurrence.intervalType = "HH" Or _
           occurrence.intervalType = "E1" Or _
           occurrence.intervalType = "E2" Then

            BuildChartExtensionTag = _
                "[EXT. REQ. FOR " & Format$(occurrence.originalWeek, "DD/MM/YYYY") & _
                "; total " & _
                Format$((occurrence.existingExtensionPercent + occurrence.extensionPercentUsed) * 100, "0") & _
                "%; imported ext " & _
                Format$(occurrence.existingExtensionAmount, "0.##") & _
                " hrs; added " & _
                Format$(occurrence.extensionPercentUsed * 100, "0") & _
                "% to Maint]"

        Else

            BuildChartExtensionTag = _
                "[EXT. REQ. FOR " & Format$(occurrence.originalWeek, "DD/MM/YYYY") & _
                "; total " & _
                Format$((occurrence.existingExtensionPercent + occurrence.extensionPercentUsed) * 100, "0") & _
                "%; imported ext " & _
                Format$(occurrence.existingExtensionAmount, "0.##") & _
                " days; added " & _
                Format$(occurrence.extensionPercentUsed * 100, "0") & _
                "% to Maint]"

        End If

    Else

        BuildChartExtensionTag = _
            "[EXT. REQ. FOR " & Format$(occurrence.originalWeek, "DD/MM/YYYY") & _
            "; " & Format$(occurrence.extensionPercentUsed * 100, "0") & _
            "% to Maint]"

    End If

End Function


'------------------------------------------------------------------------------
' Purpose : Builds a chart worksheet cell key from row and column.
'------------------------------------------------------------------------------
Public Function BuildChartCellKey(ByVal chartRow As Long, _
                                  ByVal chartColumn As Long) As String

    BuildChartCellKey = CStr(chartRow) & "," & CStr(chartColumn)

End Function


'------------------------------------------------------------------------------
' Purpose : Builds the period key used to de-duplicate chart task comments.
' Notes : Tasks in the same maintenance/down block are treated as one period.
'------------------------------------------------------------------------------
Private Function BuildChartPeriodKey(ByVal chartRow As Long, _
                                     ByVal chartColumn As Long, _
                                     ByVal scheduledWeek As Date, _
                                     ByVal displayStart As Date, _
                                     ByRef settings As PlannerSettings, _
                                     ByRef weekCyclePositions() As Long) As String

    Dim chartWeekIndex As Long
    chartWeekIndex = CLng((CLng(scheduledWeek) - CLng(displayStart)) / 7) + 1

    If chartWeekIndex < 1 Or chartWeekIndex > settings.displayWeeksShown Then
        BuildChartPeriodKey = CStr(chartRow) & "|CELL|" & CStr(chartColumn)
        Exit Function
    End If

    If chartWeekIndex > settings.WeeksBeforeForecastStart And _
       weekCyclePositions(chartWeekIndex) > settings.flyingWeeks Then

        Dim blockStartIndex As Long
        Dim blockEndIndex As Long

        blockStartIndex = FindMaintenanceBlockStart( _
            chartWeekIndex, _
            settings, _
            weekCyclePositions)

        blockEndIndex = FindMaintenanceBlockEnd( _
            chartWeekIndex, _
            settings, _
            weekCyclePositions)

        BuildChartPeriodKey = CStr(chartRow) & "|DOWN|" & _
                              CStr(blockStartIndex) & "-" & CStr(blockEndIndex)

    Else

        BuildChartPeriodKey = CStr(chartRow) & "|CELL|" & CStr(chartColumn)

    End If

End Function


'------------------------------------------------------------------------------
' Purpose : Builds unique task key for chart duplicate suppression.
'------------------------------------------------------------------------------
Private Function BuildChartTaskDeduplicationKey(ByVal chartPeriodKey As String, _
                                                ByRef occurrence As TaskOccurrence) As String

    BuildChartTaskDeduplicationKey = chartPeriodKey & "|" & _
                                     UCase$(Trim$(occurrence.taskCode)) & "|" & _
                                     UCase$(Trim$(occurrence.serialNumber)) & "|" & _
                                     UCase$(Trim$(occurrence.sequenceNumber))

End Function


'------------------------------------------------------------------------------
' Purpose : Finds start index of a contiguous maintenance block.
'------------------------------------------------------------------------------
Private Function FindMaintenanceBlockStart(ByVal weekIndex As Long, _
                                           ByRef settings As PlannerSettings, _
                                           ByRef weekCyclePositions() As Long) As Long

    Dim searchIndex As Long

    FindMaintenanceBlockStart = weekIndex

    For searchIndex = weekIndex To 1 Step -1

        If searchIndex > settings.WeeksBeforeForecastStart And _
           weekCyclePositions(searchIndex) > settings.flyingWeeks Then
            FindMaintenanceBlockStart = searchIndex
        Else
            Exit For
        End If

    Next searchIndex

End Function


'------------------------------------------------------------------------------
' Purpose : Finds end index of a contiguous maintenance block.
'------------------------------------------------------------------------------
Private Function FindMaintenanceBlockEnd(ByVal weekIndex As Long, _
                                         ByRef settings As PlannerSettings, _
                                         ByRef weekCyclePositions() As Long) As Long

    Dim searchIndex As Long

    FindMaintenanceBlockEnd = weekIndex

    For searchIndex = weekIndex To settings.displayWeeksShown

        If searchIndex > settings.WeeksBeforeForecastStart And _
           weekCyclePositions(searchIndex) > settings.flyingWeeks Then
            FindMaintenanceBlockEnd = searchIndex
        Else
            Exit For
        End If

    Next searchIndex

End Function


'------------------------------------------------------------------------------
' Purpose : Writes aggregated chart cell counts, colours, borders, and comments
' to the aircraft chart worksheet.
' Input : chartWs - aircraft chart worksheet.
' aggregation - aggregated chart cell data.
'------------------------------------------------------------------------------
Public Sub WriteChartCellAggregation(ByVal chartWs As Worksheet, _
                                     ByRef aggregation As ChartCellAggregation)

    Dim cellKey As Variant

    For Each cellKey In aggregation.CellCounts.Keys

        Dim keyParts() As String
        keyParts = Split(CStr(cellKey), ",")

        Dim chartRow As Long
        Dim chartColumn As Long

        chartRow = CLng(keyParts(0))
        chartColumn = CLng(keyParts(1))

        WriteOneAggregatedChartCell _
            chartWs, _
            chartRow, _
            chartColumn, _
            CStr(cellKey), _
            aggregation

    Next cellKey

End Sub


'------------------------------------------------------------------------------
' Purpose : Writes one aggregated chart cell.
'------------------------------------------------------------------------------
Private Sub WriteOneAggregatedChartCell(ByVal chartWs As Worksheet, _
                                        ByVal chartRow As Long, _
                                        ByVal chartColumn As Long, _
                                        ByVal cellKey As String, _
                                        ByRef aggregation As ChartCellAggregation)

    With chartWs.Cells(chartRow, chartColumn)
        .Value = aggregation.CellCounts(cellKey)
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ApplyAggregatedChartCellBorder _
        chartWs.Cells(chartRow, chartColumn), _
        cellKey, _
        aggregation

    ApplyAggregatedChartCellFill _
        chartWs, _
        chartRow, _
        chartColumn, _
        cellKey, _
        aggregation

    Dim taskText As String
    taskText = CapCommentLength( _
        aggregation.CellCounts(cellKey) & " task(s):" & vbLf & _
        aggregation.CellTasks(cellKey))

    SafeAddComment chartWs.Cells(chartRow, chartColumn), taskText

End Sub


'------------------------------------------------------------------------------
' Purpose : Applies highlighted-task border to one chart cell where required.
'------------------------------------------------------------------------------
Private Sub ApplyAggregatedChartCellBorder(ByVal targetCell As Range, _
                                           ByVal cellKey As String, _
                                           ByRef aggregation As ChartCellAggregation)

    Dim isHighlightedCell As Boolean
    isHighlightedCell = False

    If aggregation.CellHighlighted.Exists(cellKey) Then
        isHighlightedCell = CBool(aggregation.CellHighlighted(cellKey))
    End If

    If isHighlightedCell Then
        With targetCell.Borders
            .LineStyle = xlContinuous
            .Color = RGB(192, 0, 0)
            .Weight = xlThick
        End With
    End If

End Sub


'------------------------------------------------------------------------------
' Purpose : Applies fill colour to one chart cell.
'------------------------------------------------------------------------------
Private Sub ApplyAggregatedChartCellFill(ByVal chartWs As Worksheet, _
                                         ByVal chartRow As Long, _
                                         ByVal chartColumn As Long, _
                                         ByVal cellKey As String, _
                                         ByRef aggregation As ChartCellAggregation)

    Dim isPushed As Boolean
    isPushed = False

    If aggregation.CellPushed.Exists(cellKey) Then
        isPushed = CBool(aggregation.CellPushed(cellKey))
    End If

    If isPushed Then
        chartWs.Cells(chartRow, chartColumn).Interior.Color = CLR_AUTOPUSH
    Else
        chartWs.Cells(chartRow, chartColumn).Interior.Color = _
            TypeColour(CStr(chartWs.Cells(chartRow, 1).Value))
    End If

End Sub


'------------------------------------------------------------------------------
' Purpose : Merges populated chart cells across contiguous maintenance/down weeks.
' Input : chartWs - aircraft chart worksheet.
' chartTaskRowCount - number of task rows on the chart.
' settings - planner settings.
' weekCyclePositions - weekly cycle positions for this aircraft.
' aggregation - chart cell aggregation result.
'------------------------------------------------------------------------------
Public Sub MergeChartMaintenanceCells(ByVal chartWs As Worksheet, _
                                      ByVal chartTaskRowCount As Long, _
                                      ByRef settings As PlannerSettings, _
                                      ByRef weekCyclePositions() As Long, _
                                      ByRef aggregation As ChartCellAggregation)

    If chartTaskRowCount <= 0 Then Exit Sub

    Dim chartRow As Long

    For chartRow = 5 To chartTaskRowCount + 4

        Dim chartColumn As Long
        chartColumn = 4

        Do While chartColumn <= settings.displayWeeksShown + 3

            Dim weekIndex As Long
            weekIndex = chartColumn - 3

            If weekIndex < 1 Or weekIndex > settings.displayWeeksShown Then
                chartColumn = chartColumn + 1
                GoTo NextChartColumn
            End If

            If weekIndex > settings.WeeksBeforeForecastStart And _
               weekCyclePositions(weekIndex) > settings.flyingWeeks Then

                Dim mergeEndColumn As Long
                mergeEndColumn = FindMaintenanceMergeEndColumn( _
                    chartColumn, _
                    settings, _
                    weekCyclePositions)

                If mergeEndColumn > chartColumn Then
                    MergeOneChartMaintenanceBlock _
                        chartWs, _
                        chartRow, _
                        chartColumn, _
                        mergeEndColumn, _
                        aggregation
                End If

                chartColumn = mergeEndColumn + 1

            Else

                chartColumn = chartColumn + 1

            End If

NextChartColumn:
        Loop

    Next chartRow

End Sub


'------------------------------------------------------------------------------
' Purpose : Finds the final chart column in a contiguous maintenance/down block.
'------------------------------------------------------------------------------
Private Function FindMaintenanceMergeEndColumn(ByVal startColumn As Long, _
                                               ByRef settings As PlannerSettings, _
                                               ByRef weekCyclePositions() As Long) As Long

    Dim currentColumn As Long

    FindMaintenanceMergeEndColumn = startColumn

    For currentColumn = startColumn + 1 To settings.displayWeeksShown + 3

        Dim weekIndex As Long
        weekIndex = currentColumn - 3

        If weekIndex > settings.displayWeeksShown Then Exit For

        If weekCyclePositions(weekIndex) > settings.flyingWeeks Then
            FindMaintenanceMergeEndColumn = currentColumn
        Else
            Exit For
        End If

    Next currentColumn

End Function


'------------------------------------------------------------------------------
' Purpose : Merges one populated chart row across a maintenance/down block.
'------------------------------------------------------------------------------
Private Sub MergeOneChartMaintenanceBlock(ByVal chartWs As Worksheet, _
                                          ByVal chartRow As Long, _
                                          ByVal startColumn As Long, _
                                          ByVal endColumn As Long, _
                                          ByRef aggregation As ChartCellAggregation)

    If Not MaintenanceBlockHasValue(chartWs, chartRow, startColumn, endColumn) Then Exit Sub

    Dim mergedTaskCount As Long
    Dim mergedCommentText As String
    Dim mergedFillColour As Long
    Dim hasFillColour As Boolean
    Dim hasHighlightedTask As Boolean

    CollectMaintenanceBlockCellData _
        chartWs, _
        chartRow, _
        startColumn, _
        endColumn, _
        aggregation, _
        mergedTaskCount, _
        mergedCommentText, _
        mergedFillColour, _
        hasFillColour, _
        hasHighlightedTask

    ClearMaintenanceBlockCells chartWs, chartRow, startColumn, endColumn

    chartWs.Range(chartWs.Cells(chartRow, startColumn), _
                  chartWs.Cells(chartRow, endColumn)).Merge

    With chartWs.Cells(chartRow, startColumn)
        .Value = mergedTaskCount
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        If hasFillColour Then
            .Interior.Color = mergedFillColour
        End If
    End With

    If hasHighlightedTask Then
        With chartWs.Range(chartWs.Cells(chartRow, startColumn), _
                           chartWs.Cells(chartRow, endColumn)).Borders
            .LineStyle = xlContinuous
            .Color = RGB(192, 0, 0)
            .Weight = xlThick
        End With
    End If

    If Len(mergedCommentText) > 0 Then
        SafeAddComment _
            chartWs.Cells(chartRow, startColumn), _
            CapCommentLength(CStr(mergedTaskCount) & " task(s):" & vbLf & mergedCommentText)
    End If

End Sub


'------------------------------------------------------------------------------
' Purpose : Returns True if any cell in the maintenance block contains a value.
'------------------------------------------------------------------------------
Private Function MaintenanceBlockHasValue(ByVal chartWs As Worksheet, _
                                          ByVal chartRow As Long, _
                                          ByVal startColumn As Long, _
                                          ByVal endColumn As Long) As Boolean

    Dim chartColumn As Long

    For chartColumn = startColumn To endColumn
        If Len(CStr(chartWs.Cells(chartRow, chartColumn).Value)) > 0 Then
            MaintenanceBlockHasValue = True
            Exit Function
        End If
    Next chartColumn

End Function


'------------------------------------------------------------------------------
' Purpose : Collects count, comment, colour, and highlight data before merge.
'------------------------------------------------------------------------------
Private Sub CollectMaintenanceBlockCellData(ByVal chartWs As Worksheet, _
                                            ByVal chartRow As Long, _
                                            ByVal startColumn As Long, _
                                            ByVal endColumn As Long, _
                                            ByRef aggregation As ChartCellAggregation, _
                                            ByRef mergedTaskCount As Long, _
                                            ByRef mergedCommentText As String, _
                                            ByRef mergedFillColour As Long, _
                                            ByRef hasFillColour As Boolean, _
                                            ByRef hasHighlightedTask As Boolean)

    Dim chartColumn As Long

    For chartColumn = startColumn To endColumn

        If Len(CStr(chartWs.Cells(chartRow, chartColumn).Value)) > 0 Then

            mergedTaskCount = mergedTaskCount + CLng(chartWs.Cells(chartRow, chartColumn).Value)

            If Not hasFillColour Then
                mergedFillColour = chartWs.Cells(chartRow, chartColumn).Interior.Color
                hasFillColour = True
            End If

        End If

        Dim cellKey As String
        cellKey = BuildChartCellKey(chartRow, chartColumn)

        If aggregation.CellHighlighted.Exists(cellKey) Then
            If CBool(aggregation.CellHighlighted(cellKey)) Then
                hasHighlightedTask = True
            End If
        End If

        If Not chartWs.Cells(chartRow, chartColumn).Comment Is Nothing Then

            If Len(mergedCommentText) > 0 Then
                mergedCommentText = mergedCommentText & vbLf & String$(40, "-") & vbLf
            End If

            mergedCommentText = mergedCommentText & _
                RemoveChartCommentHeader( _
                    chartWs.Cells(chartRow, chartColumn).Comment.Text, _
                    CStr(chartWs.Cells(chartRow, chartColumn).Value))

            chartWs.Cells(chartRow, chartColumn).Comment.Delete

        End If

    Next chartColumn

End Sub


'------------------------------------------------------------------------------
' Purpose : Clears contents and fills from a maintenance block before merging.
'------------------------------------------------------------------------------
Private Sub ClearMaintenanceBlockCells(ByVal chartWs As Worksheet, _
                                       ByVal chartRow As Long, _
                                       ByVal startColumn As Long, _
                                       ByVal endColumn As Long)

    Dim chartColumn As Long

    For chartColumn = startColumn To endColumn
        chartWs.Cells(chartRow, chartColumn).ClearContents
        chartWs.Cells(chartRow, chartColumn).Interior.ColorIndex = xlNone
    Next chartColumn

End Sub


'------------------------------------------------------------------------------
' Purpose : Removes the "x task(s):" header from an existing chart comment.
'------------------------------------------------------------------------------
Private Function RemoveChartCommentHeader(ByVal commentText As String, _
                                          ByVal taskCountText As String) As String

    Dim headerText As String
    headerText = taskCountText & " task(s):" & vbLf

    RemoveChartCommentHeader = Replace(commentText, headerText, vbNullString)

End Function


'------------------------------------------------------------------------------
' Purpose : Marks original due-week cells where an extension must be raised.
' Input : chartWs - aircraft chart worksheet.
' taskOccurrences - populated task occurrence records.
' totalOccurrences - number of populated occurrences.
' rowLookup - chart row lookup by interval pair key.
' columnLookup - chart column lookup by week start date.
' Notes : Called before maintenance-cell merging so flying-week cells remain separate.
'------------------------------------------------------------------------------
Public Sub WriteExtensionDueIndicators(ByVal chartWs As Worksheet, _
                                       ByRef taskOccurrences() As TaskOccurrence, _
                                       ByVal totalOccurrences As Long, _
                                       ByVal rowLookup As Object, _
                                       ByVal columnLookup As Object)

    Dim occurrenceIndex As Long

    For occurrenceIndex = 1 To totalOccurrences

        Dim occurrence As TaskOccurrence
        occurrence = taskOccurrences(occurrenceIndex)

        If Not occurrence.wasAutoPushed Then GoTo NextOccurrence
        If CLng(occurrence.originalWeek) = CLng(occurrence.scheduledWeek) Then GoTo NextOccurrence

        Dim pairKey As String
        pairKey = BuildChartPairKey(occurrence.intervalType, occurrence.intervalValue)

        If Not rowLookup.Exists(pairKey) Then GoTo NextOccurrence

        Dim chartRow As Long
        chartRow = CLng(rowLookup(pairKey))

        Dim originalWeekKey As Long
        originalWeekKey = CLng(occurrence.originalWeek)

        If Not columnLookup.Exists(originalWeekKey) Then GoTo NextOccurrence

        Dim chartColumn As Long
        chartColumn = CLng(columnLookup(originalWeekKey))

        MarkExtensionDueChartCell chartWs, chartRow, chartColumn, occurrence

NextOccurrence:
    Next occurrenceIndex

End Sub


'------------------------------------------------------------------------------
' Purpose : Writes or appends an extension-due indicator to one chart cell.
'------------------------------------------------------------------------------
Private Sub MarkExtensionDueChartCell(ByVal chartWs As Worksheet, _
                                      ByVal chartRow As Long, _
                                      ByVal chartColumn As Long, _
                                      ByRef occurrence As TaskOccurrence)

    Dim targetCell As Range
    Set targetCell = chartWs.Cells(chartRow, chartColumn)

    If Len(Trim$(CStr(targetCell.Value))) > 0 Then Exit Sub

    With targetCell
        .Value = "EXT"
        .Font.Name = "Arial"
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = RGB(192, 64, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = CLR_EXT_DUE
    End With

    Dim indicatorText As String
    indicatorText = "Extension must be raised for original due week " & _
                    Format$(occurrence.originalWeek, "DD/MM/YYYY") & vbLf & _
                    occurrence.taskCode & " - " & TruncateDesc(occurrence.taskDescription) & _
                    RightAlignTag(BuildChartExtensionTag(occurrence))

    SafeAddComment targetCell, CapCommentLength(indicatorText)

End Sub


