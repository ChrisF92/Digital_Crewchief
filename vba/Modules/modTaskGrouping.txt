Attribute VB_Name = "modTaskGrouping"
Option Explicit

'------------------------------------------------------------------------------
' Module : modTaskGrouping
' Purpose : Groups task occurrences that are close enough to be aligned.
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Purpose : Applies grouping rules to task occurrences for one aircraft.
' Input : taskOccurrences - populated task occurrence records.
' totalOccurrences - number of populated records.
' groupingTolerance - tolerance percentage as decimal, e.g. 0.1.
' settings - planner settings.
' weekHHRates - weekly HH rates.
' weekE1Rates - weekly E1 rates.
' weekE2Rates - weekly E2 rates.
'------------------------------------------------------------------------------
Public Sub ApplyTaskGrouping(ByRef taskOccurrences() As TaskOccurrence, _
                             ByVal totalOccurrences As Long, _
                             ByVal groupingTolerance As Double, _
                             ByRef settings As PlannerSettings, _
                             ByRef weekHHRates() As Double, _
                             ByRef weekE1Rates() As Double, _
                             ByRef weekE2Rates() As Double)

    If groupingTolerance <= 0 Then Exit Sub
    If totalOccurrences <= 1 Then Exit Sub

    Dim groupDictionary As Object
    Set groupDictionary = BuildTaskGroupingDictionary(taskOccurrences, totalOccurrences)

    Dim groupKey As Variant

    For Each groupKey In groupDictionary.Keys

        ApplyGroupingForOneKey _
            taskOccurrences, _
            totalOccurrences, _
            CStr(groupKey), _
            CDbl(groupDictionary(groupKey)), _
            groupingTolerance, _
            settings, _
            weekHHRates, _
            weekE1Rates, _
            weekE2Rates

    Next groupKey

End Sub


'------------------------------------------------------------------------------
' Purpose : Builds unique grouping keys from task interval type and value.
'------------------------------------------------------------------------------
Private Function BuildTaskGroupingDictionary(ByRef taskOccurrences() As TaskOccurrence, _
                                             ByVal totalOccurrences As Long) As Object

    Dim groupDictionary As Object
    Set groupDictionary = CreateObject("Scripting.Dictionary")

    Dim occurrenceIndex As Long

    For occurrenceIndex = 1 To totalOccurrences

        Dim groupKey As String
        groupKey = BuildTaskGroupingKey(taskOccurrences(occurrenceIndex))

        If Not groupDictionary.Exists(groupKey) Then
            groupDictionary.Add groupKey, taskOccurrences(occurrenceIndex).intervalValue
        End If

    Next occurrenceIndex

    Set BuildTaskGroupingDictionary = groupDictionary

End Function


'------------------------------------------------------------------------------
' Purpose : Builds the grouping key for one occurrence.
'------------------------------------------------------------------------------
Private Function BuildTaskGroupingKey(ByRef occurrence As TaskOccurrence) As String

    BuildTaskGroupingKey = occurrence.intervalType & "|" & CStr(occurrence.intervalValue)

End Function


'------------------------------------------------------------------------------
' Purpose : Applies grouping for one interval type/value group.
'------------------------------------------------------------------------------
Private Sub ApplyGroupingForOneKey(ByRef taskOccurrences() As TaskOccurrence, _
                                   ByVal totalOccurrences As Long, _
                                   ByVal groupKey As String, _
                                   ByVal groupInterval As Double, _
                                   ByVal groupingTolerance As Double, _
                                   ByRef settings As PlannerSettings, _
                                   ByRef weekHHRates() As Double, _
                                   ByRef weekE1Rates() As Double, _
                                   ByRef weekE2Rates() As Double)

    Dim groupIntervalType As String
    groupIntervalType = Left$(groupKey, InStr(groupKey, "|") - 1)

    Dim toleranceDays As Long
    toleranceDays = CalculateGroupingToleranceDays( _
        groupIntervalType, _
        groupInterval, _
        groupingTolerance, _
        settings, _
        weekHHRates, _
        weekE1Rates, _
        weekE2Rates)

    Dim groupedIndexes() As Long
    Dim groupedCount As Long

    groupedIndexes = GetGroupableOccurrenceIndexes( _
        taskOccurrences, _
        totalOccurrences, _
        groupKey, _
        groupedCount)

    If groupedCount <= 1 Then Exit Sub

    AlignGroupedOccurrences _
        taskOccurrences, _
        groupedIndexes, _
        groupedCount, _
        toleranceDays

End Sub


'------------------------------------------------------------------------------
' Purpose : Calculates the date tolerance used for grouping.
'------------------------------------------------------------------------------
Private Function CalculateGroupingToleranceDays(ByVal intervalType As String, _
                                                ByVal intervalValue As Double, _
                                                ByVal groupingTolerance As Double, _
                                                ByRef settings As PlannerSettings, _
                                                ByRef weekHHRates() As Double, _
                                                ByRef weekE1Rates() As Double, _
                                                ByRef weekE2Rates() As Double) As Long

    Dim toleranceAmount As Double
    toleranceAmount = intervalValue * groupingTolerance

    If intervalType = "DD" Then

        CalculateGroupingToleranceDays = CLng(toleranceAmount)

    Else

        Dim averageWeeklyRate As Double
        averageWeeklyRate = CalculateAverageWeeklyRate( _
            intervalType, _
            settings.displayWeeksShown, _
            weekHHRates, _
            weekE1Rates, _
            weekE2Rates)

        If averageWeeklyRate > 0 Then
            CalculateGroupingToleranceDays = CLng((toleranceAmount / averageWeeklyRate) * 7)
        Else
            CalculateGroupingToleranceDays = 7
        End If

    End If

    If CalculateGroupingToleranceDays < 1 Then
        CalculateGroupingToleranceDays = 1
    End If

End Function


'------------------------------------------------------------------------------
' Purpose : Calculates average weekly flying rate for one interval type.
'------------------------------------------------------------------------------
Private Function CalculateAverageWeeklyRate(ByVal intervalType As String, _
                                            ByVal displayWeeksShown As Long, _
                                            ByRef weekHHRates() As Double, _
                                            ByRef weekE1Rates() As Double, _
                                            ByRef weekE2Rates() As Double) As Double

    Dim totalRate As Double
    Dim flyingWeekCount As Long
    Dim weekIndex As Long

    For weekIndex = 1 To displayWeeksShown

        Dim weeklyRate As Double

        Select Case intervalType
            Case "HH"
                weeklyRate = weekHHRates(weekIndex)

            Case "E1"
                weeklyRate = weekE1Rates(weekIndex)

            Case "E2"
                weeklyRate = weekE2Rates(weekIndex)

            Case Else
                weeklyRate = 0
        End Select

        If weeklyRate > 0 Then
            totalRate = totalRate + weeklyRate
            flyingWeekCount = flyingWeekCount + 1
        End If

    Next weekIndex

    If flyingWeekCount > 0 Then
        CalculateAverageWeeklyRate = totalRate / flyingWeekCount
    Else
        CalculateAverageWeeklyRate = 0
    End If

End Function


'------------------------------------------------------------------------------
' Purpose : Returns indexes of occurrences that can participate in grouping.
'------------------------------------------------------------------------------
Private Function GetGroupableOccurrenceIndexes(ByRef taskOccurrences() As TaskOccurrence, _
                                               ByVal totalOccurrences As Long, _
                                               ByVal groupKey As String, _
                                               ByRef groupedCount As Long) As Long()

    Dim groupedIndexes() As Long
    ReDim groupedIndexes(1 To totalOccurrences)

    groupedCount = 0

    Dim occurrenceIndex As Long

    For occurrenceIndex = 1 To totalOccurrences

        If BuildTaskGroupingKey(taskOccurrences(occurrenceIndex)) = groupKey Then

            If Not IsNoPullTaskCode(taskOccurrences(occurrenceIndex).taskCode) Then
                groupedCount = groupedCount + 1
                groupedIndexes(groupedCount) = occurrenceIndex
            End If

        End If

    Next occurrenceIndex

    GetGroupableOccurrenceIndexes = groupedIndexes

End Function


'------------------------------------------------------------------------------
' Purpose : Moves grouped occurrences to the earliest nearby occurrence date.
'------------------------------------------------------------------------------
Private Sub AlignGroupedOccurrences(ByRef taskOccurrences() As TaskOccurrence, _
                                    ByRef groupedIndexes() As Long, _
                                    ByVal groupedCount As Long, _
                                    ByVal toleranceDays As Long)

    Dim sourceIndex As Long

    For sourceIndex = 1 To groupedCount

        Dim occurrenceIndex As Long
        occurrenceIndex = groupedIndexes(sourceIndex)

        Dim baseWeek As Date
        baseWeek = taskOccurrences(occurrenceIndex).scheduledWeek

        Dim earliestGroupedWeek As Date
        earliestGroupedWeek = FindEarliestWeekWithinTolerance( _
            taskOccurrences, _
            groupedIndexes, _
            groupedCount, _
            sourceIndex, _
            baseWeek, _
            toleranceDays)

        If earliestGroupedWeek <> baseWeek Then

            With taskOccurrences(occurrenceIndex)

                .scheduledWeek = earliestGroupedWeek

                If earliestGroupedWeek < .originalWeek Then
                    .GroupedDirection = -1
                Else
                    .GroupedDirection = 1
                End If

            End With

        End If

    Next sourceIndex

End Sub


'------------------------------------------------------------------------------
' Purpose : Finds the earliest occurrence date within the grouping tolerance.
'------------------------------------------------------------------------------
Private Function FindEarliestWeekWithinTolerance(ByRef taskOccurrences() As TaskOccurrence, _
                                                 ByRef groupedIndexes() As Long, _
                                                 ByVal groupedCount As Long, _
                                                 ByVal sourceIndex As Long, _
                                                 ByVal baseWeek As Date, _
                                                 ByVal toleranceDays As Long) As Date

    Dim earliestWeek As Date
    earliestWeek = baseWeek

    Dim compareIndex As Long

    For compareIndex = 1 To groupedCount

        If compareIndex <> sourceIndex Then

            Dim comparisonWeek As Date
            comparisonWeek = taskOccurrences(groupedIndexes(compareIndex)).scheduledWeek

            If Abs(CLng(comparisonWeek) - CLng(baseWeek)) <= toleranceDays Then
                If comparisonWeek < earliestWeek Then
                    earliestWeek = comparisonWeek
                End If
            End If

        End If

    Next compareIndex

    FindEarliestWeekWithinTolerance = earliestWeek

End Function


