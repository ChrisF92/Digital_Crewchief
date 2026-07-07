Attribute VB_Name = "modImportLayout"
Option Explicit

'------------------------------------------------------------------------------
' Module : modImportLayout
' Purpose : Defines the fixed internal layout of each aircraft import sheet.
'
' Notes:
' These constants describe the destination import sheet layout, not the source
' export file. Source export files should continue to be handled by header
' names because their column order may change.
'------------------------------------------------------------------------------

Public Const CHART_SHEET_SUFFIX As String = "_CHART"
Public Const IMPORT_SHEET_SUFFIX As String = "_IMPORT"

Public Const IMPORT_FIRST_DATA_ROW As Long = 2
Public Const IMPORT_LAST_DATA_ROW As Long = 1500
Public Const IMPORT_MAX_DATA_ROWS As Long = IMPORT_LAST_DATA_ROW - IMPORT_FIRST_DATA_ROW + 1

Public Const IMPORT_COL_SHOW As Long = 1
Public Const IMPORT_COL_INTERVAL_TYPE As Long = 2
Public Const IMPORT_COL_LIFE_REMAINING As Long = 3
Public Const IMPORT_COL_TASK_CODE As Long = 4
Public Const IMPORT_COL_TASK_SEQUENCE As Long = 5
Public Const IMPORT_COL_TASK_DESCRIPTION As Long = 6
Public Const IMPORT_COL_SERIAL As Long = 7
Public Const IMPORT_COL_ACTUAL_INTERVAL As Long = 8
Public Const IMPORT_COL_EXTENSION As Long = 9
Public Const IMPORT_COL_PROJECTED_DUE_DATE As Long = 10
Public Const IMPORT_COL_END_ITEM_SERIAL As Long = 11

Public Const IMPORT_FIRST_COL As Long = IMPORT_COL_SHOW
Public Const IMPORT_LAST_COL As Long = IMPORT_COL_END_ITEM_SERIAL

Public Const IMPORT_LAST_IMPORT_LABEL_CELL As String = "M1"
Public Const IMPORT_LAST_IMPORT_VALUE_CELL As String = "M2"


'------------------------------------------------------------------------------
' Purpose : Returns the full data range for an aircraft import sheet.
' Input : importWs - aircraft import worksheet.
' Output : Range covering the internal import table data rows and columns.
'------------------------------------------------------------------------------
Public Function GetImportDataRange(ByVal importWs As Worksheet) As Range

    Set GetImportDataRange = importWs.Range( _
        importWs.Cells(IMPORT_FIRST_DATA_ROW, IMPORT_FIRST_COL), _
        importWs.Cells(IMPORT_LAST_DATA_ROW, IMPORT_LAST_COL))

End Function


'------------------------------------------------------------------------------
' Purpose : Returns True if an import row contains an interval type.
' Input : importWs - aircraft import worksheet.
' importRow - row number to inspect.
' Output : True when the row appears to contain task data.
'------------------------------------------------------------------------------
Public Function ImportRowHasTaskData(ByVal importWs As Worksheet, _
                                     ByVal importRow As Long) As Boolean

    ImportRowHasTaskData = _
        Len(Trim$(CStr(importWs.Cells(importRow, IMPORT_COL_INTERVAL_TYPE).Value))) > 0

End Function


'------------------------------------------------------------------------------
' Purpose : Clears the task data from an aircraft import sheet.
' Input : importWs - aircraft import worksheet.
'------------------------------------------------------------------------------
Public Sub ClearImportData(ByVal importWs As Worksheet)

    GetImportDataRange(importWs).ClearContents

End Sub


'------------------------------------------------------------------------------
' Purpose : Stores the last import timestamp on the aircraft import sheet.
' Input : importWs - aircraft import worksheet.
'------------------------------------------------------------------------------
Public Sub StampLastImportTime(ByVal importWs As Worksheet)

    importWs.Range(IMPORT_LAST_IMPORT_LABEL_CELL).Value = "Last Import"
    importWs.Range(IMPORT_LAST_IMPORT_VALUE_CELL).Value = Now
    importWs.Range(IMPORT_LAST_IMPORT_VALUE_CELL).NumberFormat = "DD/MM/YYYY HH:MM"

End Sub


'------------------------------------------------------------------------------
' Purpose : Returns the stored last-import timestamp for an aircraft import sheet.
' Input : importWs - aircraft import worksheet.
' Output : Last import date/time, or Empty when never imported.
'------------------------------------------------------------------------------
Public Function GetLastImportTime(ByVal importWs As Worksheet) As Variant

    GetLastImportTime = importWs.Range(IMPORT_LAST_IMPORT_VALUE_CELL).Value

End Function


'------------------------------------------------------------------------------
' Purpose : Returns True when the last import falls in the current calendar week.
' Notes : Weeks run Monday to Sunday, matching planner calendar conventions.
' Input : importWs - aircraft import worksheet.
' Output : True when import is current; False when missing or from another week.
'------------------------------------------------------------------------------
Public Function IsImportCurrentWeek(ByVal importWs As Worksheet) As Boolean

    Dim lastImport As Variant
    lastImport = GetLastImportTime(importWs)

    If IsEmpty(lastImport) Then Exit Function
    If Not IsDate(lastImport) Then Exit Function

    IsImportCurrentWeek = (GetWeekMonday(CDate(lastImport)) = GetCurrentWeekMonday())

End Function


'------------------------------------------------------------------------------
' Purpose : Counts roster aircraft whose SITS import is missing or from another week.
' Input : tailNumbers - active aircraft tail numbers.
' aircraftCount - number of aircraft in the roster.
' Output : Count of aircraft needing re-import.
'------------------------------------------------------------------------------
Public Function CountOutOfDateImports(ByRef tailNumbers() As String, _
                                      ByVal aircraftCount As Long) As Long

    Dim aircraftIndex As Long
    Dim importWs As Worksheet

    For aircraftIndex = 1 To aircraftCount

        If SheetExists(tailNumbers(aircraftIndex) & IMPORT_SHEET_SUFFIX) Then

            Set importWs = ThisWorkbook.Worksheets(tailNumbers(aircraftIndex) & IMPORT_SHEET_SUFFIX)

            If Not IsImportCurrentWeek(importWs) Then
                CountOutOfDateImports = CountOutOfDateImports + 1
            End If

        Else
            CountOutOfDateImports = CountOutOfDateImports + 1
        End If

    Next aircraftIndex

End Function


