Attribute VB_Name = "modTaskOccurrenceFormatting"
Option Explicit

'------------------------------------------------------------------------------
' Module : modTaskOccurrenceFormatting
' Purpose : Formats TaskOccurrence records for dashboard and chart comments.
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Purpose : Builds a short interval label such as "HH 25", "E1 50", or "DD 180".
' Input : occurrence - task occurrence to format.
' Output : Display label for dashboard/chart summary buckets.
'------------------------------------------------------------------------------
Public Function BuildOccurrenceIntervalLabel(ByRef occurrence As TaskOccurrence) As String

    If occurrence.intervalValue = Int(occurrence.intervalValue) Then
        BuildOccurrenceIntervalLabel = occurrence.intervalType & " " & _
                                       CStr(CLng(occurrence.intervalValue))
    Else
        BuildOccurrenceIntervalLabel = occurrence.intervalType & " " & _
                                       Format$(occurrence.intervalValue, "0.##")
    End If

End Function


'------------------------------------------------------------------------------
' Purpose : Builds the due/life remaining text for a task occurrence.
' Input : occurrence - task occurrence to format.
' Output : Text such as "Due: 10/07/2026" or "Due: 25:30HH".
'------------------------------------------------------------------------------
Public Function BuildOccurrenceDueText(ByRef occurrence As TaskOccurrence) As String

    If occurrence.intervalType = "DD" Then
        BuildOccurrenceDueText = "Due: " & Format$(occurrence.firstDueDate, "DD/MM/YYYY")
    Else
        BuildOccurrenceDueText = "Due: " & _
                                 FormatHoursToHMM(occurrence.lifeRemaining) & _
                                 occurrence.intervalType
    End If

End Function


'------------------------------------------------------------------------------
' Purpose : Builds the extension tag for a pushed task.
' Input : occurrence - task occurrence to format.
' Output : Extension tag, or blank if the occurrence was not pushed.
'------------------------------------------------------------------------------
Public Function BuildOccurrenceExtensionTag(ByRef occurrence As TaskOccurrence) As String

    If Not occurrence.wasAutoPushed Then
        BuildOccurrenceExtensionTag = vbNullString
        Exit Function
    End If

    If occurrence.existingExtensionPercent > 0 Then

        BuildOccurrenceExtensionTag = "[EXT total " & _
            Format$( _
                (occurrence.existingExtensionPercent + occurrence.extensionPercentUsed) * 100, _
                "0") & _
            "%; added " & _
            Format$(occurrence.extensionPercentUsed * 100, "0") & _
            "%]"

    Else

        BuildOccurrenceExtensionTag = "[EXT " & _
            Format$(occurrence.extensionPercentUsed * 100, "0") & _
            "%]"

    End If

End Function


'------------------------------------------------------------------------------
' Purpose : Builds the pull-forward tag for a pulled task.
' Input : occurrence - task occurrence to format.
' Output : Pull tag, or blank if the occurrence was not pulled.
'------------------------------------------------------------------------------
Public Function BuildOccurrencePullTag(ByRef occurrence As TaskOccurrence) As String

    If occurrence.wasAutoPulled Then
        BuildOccurrencePullTag = "[PULLED " & _
                                 Format$(occurrence.pullPercentUsed * 100, "0") & _
                                 "%]"
    Else
        BuildOccurrencePullTag = vbNullString
    End If

End Function


'------------------------------------------------------------------------------
' Purpose : Builds a full dashboard detail line for one occurrence.
' Input : occurrence - task occurrence to format.
' Output : Detail line used in dashboard comments.
'------------------------------------------------------------------------------
Public Function BuildDashboardOccurrenceDetail(ByRef occurrence As TaskOccurrence) As String

    Dim detailText As String

    detailText = occurrence.taskCode & " - " & _
                 TruncateDesc(occurrence.taskDescription) & _
                 RightAlignTag("[" & BuildOccurrenceDueText(occurrence) & "]")

    Dim extensionTag As String
    extensionTag = BuildOccurrenceExtensionTag(occurrence)

    If Len(extensionTag) > 0 Then
        detailText = detailText & RightAlignTag(extensionTag)
    End If

    Dim pullTag As String
    pullTag = BuildOccurrencePullTag(occurrence)

    If Len(pullTag) > 0 Then
        detailText = detailText & RightAlignTag(pullTag)
    End If

    BuildDashboardOccurrenceDetail = detailText

End Function


'------------------------------------------------------------------------------
' Purpose : Adds one occurrence to a dashboard summary bucket if it has not
' already been added.
' Input : summaryDictionary - dictionary containing summary counts.
' detailDictionary - dictionary containing detailed task lines.
' taskCodeDictionary - dictionary used to prevent duplicate task keys.
' taskKey - unique occurrence key.
' intervalLabel - bucket label, e.g. "HH 25".
' detailText - formatted dashboard detail line.
'------------------------------------------------------------------------------
Public Sub AddOccurrenceToDashboardBucket(ByVal summaryDictionary As Object, _
                                          ByVal detailDictionary As Object, _
                                          ByVal taskCodeDictionary As Object, _
                                          ByVal taskKey As String, _
                                          ByVal intervalLabel As String, _
                                          ByVal detailText As String)

    If taskCodeDictionary.Exists(taskKey) Then Exit Sub

    AddSummaryItem summaryDictionary, detailDictionary, intervalLabel, detailText
    taskCodeDictionary.Add taskKey, True

End Sub


