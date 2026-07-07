Attribute VB_Name = "modTaskRules"
Option Explicit

'------------------------------------------------------------------------------
' Module : modTaskRules
' Purpose : Loads and matches Task_Rules worksheet prefixes during refresh.
'------------------------------------------------------------------------------

Public Const TASK_RULES_SHEET_NAME As String = "Task_Rules"
Public Const TASK_RULE_FIRST_DATA_ROW As Long = 2

Public Const TASK_RULE_COL_HIGHLIGHT As Long = 1
Public Const TASK_RULE_COL_NO_EXTEND As Long = 3
Public Const TASK_RULE_COL_NO_PULL As Long = 5

Private m_TaskRuleCache As Object
Private m_TaskRuleCacheLoaded As Boolean


'------------------------------------------------------------------------------
' Purpose : Loads Task_Rules prefixes into memory for the current refresh.
'------------------------------------------------------------------------------
Public Sub LoadTaskRuleCache()

    Set m_TaskRuleCache = CreateObject("Scripting.Dictionary")
    m_TaskRuleCacheLoaded = False

    Dim ws As Worksheet
    Set ws = RuleSheet()

    If ws Is Nothing Then
        m_TaskRuleCacheLoaded = True
        Exit Sub
    End If

    CacheTaskRuleColumn ws, TASK_RULE_COL_HIGHLIGHT
    CacheTaskRuleColumn ws, TASK_RULE_COL_NO_EXTEND
    CacheTaskRuleColumn ws, TASK_RULE_COL_NO_PULL

    m_TaskRuleCacheLoaded = True

End Sub


'------------------------------------------------------------------------------
' Purpose : Clears cached Task_Rules prefixes after refresh completes.
'------------------------------------------------------------------------------
Public Sub ClearTaskRuleCache()

    Set m_TaskRuleCache = Nothing
    m_TaskRuleCacheLoaded = False

End Sub


'------------------------------------------------------------------------------
' Purpose : Checks whether a task code should be highlighted.
' Input : taskCode - task code to test.
' Output : True if it matches the highlight rules column, otherwise False.
'------------------------------------------------------------------------------
Public Function IsHighlightedTaskCode(taskCode As String) As Boolean

    Dim tc As String
    tc = UCase(Trim(CStr(taskCode)))

    If Len(tc) = 0 Then
        IsHighlightedTaskCode = False
        Exit Function
    End If

    IsHighlightedTaskCode = TaskCodeMatchesRuleColumn(taskCode, TASK_RULE_COL_HIGHLIGHT)

End Function


'------------------------------------------------------------------------------
' Purpose : Checks whether a task code is excluded from extension logic.
' Input : taskCode - task code to test.
' Output : True if the task matches the no-extension rule column.
'------------------------------------------------------------------------------
Public Function IsNoExtendTaskCode(taskCode As String) As Boolean
    IsNoExtendTaskCode = TaskCodeMatchesRuleColumn(taskCode, TASK_RULE_COL_NO_EXTEND)
End Function


'------------------------------------------------------------------------------
' Purpose : Checks whether a task code is excluded from pull-forward logic.
' Input : taskCode - task code to test.
' Output : True if the task matches the no-pull rule column.
'------------------------------------------------------------------------------
Public Function IsNoPullTaskCode(taskCode As String) As Boolean
    IsNoPullTaskCode = TaskCodeMatchesRuleColumn(taskCode, TASK_RULE_COL_NO_PULL)
End Function


'------------------------------------------------------------------------------
' Purpose : Loads task rule prefixes from one Task_Rules column.
' Input : ruleCol - column number on Task_Rules (1, 3, or 5).
' Output : Collection of uppercase trimmed prefixes, in sheet order.
'------------------------------------------------------------------------------
Public Function LoadTaskRulePrefixes(ByVal ruleCol As Long) As Collection

    Dim ws As Worksheet
    Set ws = RuleSheet()

    If ws Is Nothing Then
        Set LoadTaskRulePrefixes = New Collection
        Exit Function
    End If

    Set LoadTaskRulePrefixes = ReadPrefixesFromColumn(ws, ruleCol)

End Function


'------------------------------------------------------------------------------
' Purpose : Writes all three Task_Rules prefix columns from memory collections.
' Input : highlightPrefixes, noExtendPrefixes, noPullPrefixes - prefix lists.
' Output : Raises an error when Task_Rules is missing or protected.
'------------------------------------------------------------------------------
Public Sub SaveAllTaskRules(ByVal highlightPrefixes As Collection, _
                            ByVal noExtendPrefixes As Collection, _
                            ByVal noPullPrefixes As Collection)

    Dim ws As Worksheet
    Set ws = RuleSheet()

    If ws Is Nothing Then
        Err.Raise vbObjectError + 2210, _
                  "SaveAllTaskRules", _
                  "Task_Rules sheet was not found."
    End If

    WritePrefixesToColumn ws, TASK_RULE_COL_HIGHLIGHT, highlightPrefixes
    WritePrefixesToColumn ws, TASK_RULE_COL_NO_EXTEND, noExtendPrefixes
    WritePrefixesToColumn ws, TASK_RULE_COL_NO_PULL, noPullPrefixes

End Sub


'------------------------------------------------------------------------------
' Purpose : Normalises a task rule prefix for storage and comparison.
' Input : rawPrefix - user-entered prefix text.
' Output : Uppercase trimmed prefix.
'------------------------------------------------------------------------------
Public Function NormalizeTaskRulePrefix(ByVal rawPrefix As String) As String

    NormalizeTaskRulePrefix = UCase$(Trim$(CStr(rawPrefix)))

End Function


'------------------------------------------------------------------------------
' Purpose : Returns the Task_Rules worksheet if it exists.
' Output : Task_Rules worksheet object, or Nothing if missing.
'------------------------------------------------------------------------------
Private Function RuleSheet() As Worksheet

    Set RuleSheet = Nothing

    On Error Resume Next
    Set RuleSheet = ThisWorkbook.Sheets(TASK_RULES_SHEET_NAME)
    On Error GoTo 0

End Function


'------------------------------------------------------------------------------
' Purpose : Stores one Task_Rules column as a collection of uppercase prefixes.
'------------------------------------------------------------------------------
Private Sub CacheTaskRuleColumn(ByVal ruleWs As Worksheet, ByVal ruleCol As Long)

    Set m_TaskRuleCache(CStr(ruleCol)) = ReadPrefixesFromColumn(ruleWs, ruleCol)

End Sub


'------------------------------------------------------------------------------
' Purpose : Reads non-blank prefixes from one Task_Rules column.
'------------------------------------------------------------------------------
Private Function ReadPrefixesFromColumn(ByVal ruleWs As Worksheet, _
                                        ByVal ruleCol As Long) As Collection

    Dim prefixes As Collection
    Set prefixes = New Collection

    Dim lastRow As Long
    lastRow = ruleWs.Cells(ruleWs.Rows.Count, ruleCol).End(xlUp).Row

    Dim rowNumber As Long
    Dim listedCode As String

    For rowNumber = TASK_RULE_FIRST_DATA_ROW To lastRow

        listedCode = NormalizeTaskRulePrefix(CStr(ruleWs.Cells(rowNumber, ruleCol).Value))

        If Len(listedCode) > 0 Then
            prefixes.Add listedCode
        End If

    Next rowNumber

    Set ReadPrefixesFromColumn = prefixes

End Function


'------------------------------------------------------------------------------
' Purpose : Replaces one Task_Rules column with a prefix collection.
'------------------------------------------------------------------------------
Private Sub WritePrefixesToColumn(ByVal ruleWs As Worksheet, _
                                  ByVal ruleCol As Long, _
                                  ByVal prefixes As Collection)

    Dim lastRow As Long
    lastRow = ruleWs.Cells(ruleWs.Rows.Count, ruleCol).End(xlUp).Row

    If lastRow < TASK_RULE_FIRST_DATA_ROW Then
        lastRow = TASK_RULE_FIRST_DATA_ROW
    End If

    ruleWs.Range( _
        ruleWs.Cells(TASK_RULE_FIRST_DATA_ROW, ruleCol), _
        ruleWs.Cells(lastRow, ruleCol)).ClearContents

    Dim prefixIndex As Long

    For prefixIndex = 1 To prefixes.Count
        ruleWs.Cells(TASK_RULE_FIRST_DATA_ROW + prefixIndex - 1, ruleCol).Value = prefixes(prefixIndex)
    Next prefixIndex

End Sub


'------------------------------------------------------------------------------
' Purpose : Checks whether a task code matches any rule prefix in a specified column.
' Input : taskCode - task code to test.
' ruleCol - Task_Rules column containing rule prefixes.
' Output : True if the task code starts with any listed rule prefix.
'------------------------------------------------------------------------------
Private Function TaskCodeMatchesRuleColumn(taskCode As String, ruleCol As Long) As Boolean

    Dim tc As String
    tc = UCase$(Trim$(CStr(taskCode)))

    If Len(tc) = 0 Then
        TaskCodeMatchesRuleColumn = False
        Exit Function
    End If

    If Not m_TaskRuleCacheLoaded Or m_TaskRuleCache Is Nothing Then
        TaskCodeMatchesRuleColumn = False
        Exit Function
    End If

    If Not m_TaskRuleCache.Exists(CStr(ruleCol)) Then
        TaskCodeMatchesRuleColumn = False
        Exit Function
    End If

    Dim prefixes As Collection
    Set prefixes = m_TaskRuleCache(CStr(ruleCol))

    Dim prefixIndex As Long
    Dim listedCode As String

    For prefixIndex = 1 To prefixes.Count

        listedCode = prefixes(prefixIndex)

        If Left$(tc, Len(listedCode)) = listedCode Then
            TaskCodeMatchesRuleColumn = True
            Exit Function
        End If

    Next prefixIndex

    TaskCodeMatchesRuleColumn = False

End Function
