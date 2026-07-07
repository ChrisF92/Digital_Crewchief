Attribute VB_Name = "modSheetCreation"
Option Explicit

'------------------------------------------------------------------------------
' Module : modSheetCreation
' Purpose : Creates and formats aircraft import/chart sheets and chart navigation controls.
'------------------------------------------------------------------------------

'------------------------------------------------------------------------------
' Purpose : Creates and formats a new aircraft import worksheet.
' Input : tail - aircraft tail number.
' Output : Newly created import worksheet.
'------------------------------------------------------------------------------
Public Function CreateImportSheet(tail As String) As Worksheet
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))

    ws.Name = tail & IMPORT_SHEET_SUFFIX
    ws.Tab.Color = RGB(128, 128, 128)

    Dim h As Variant
    h = Array("Show?", "Interval Type Code", "Life Remaining", "Task Code", "Task Sequence", _
              "Task Description", "Serial", "Actual Interval", "Extension", _
              "Projected Due Date", "End Item Serial")

    Dim i As Long

    For i = 0 To UBound(h)
        With ws.Cells(1, i + 1)
            .Value = h(i)
            .Font.Name = "Arial"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(47, 84, 150)
            .HorizontalAlignment = xlCenter
        End With
    Next i

    Dim ww As Variant
    ww = Array(8, 18, 14, 22, 14, 50, 14, 14, 12, 18, 16)

    For i = 0 To UBound(ww)
        ws.Columns(i + 1).ColumnWidth = ww(i)
    Next i

    Dim r As Long

    For r = IMPORT_FIRST_DATA_ROW To IMPORT_LAST_DATA_ROW
        With ws.Cells(r, 1)
            .Value = "Y"
            .Font.Name = "Arial"
            .Font.Size = 10
            .Font.Color = RGB(0, 0, 255)
            .HorizontalAlignment = xlCenter
        End With
    Next r

    With ws.Range(ws.Cells(IMPORT_FIRST_DATA_ROW, IMPORT_COL_SHOW), ws.Cells(IMPORT_LAST_DATA_ROW, IMPORT_COL_SHOW)).Validation
        .Delete
        .Add xlValidateList, xlValidAlertStop, , "Y,N"
        .ShowError = True
    End With

    ws.Range(ws.Cells(IMPORT_FIRST_DATA_ROW, IMPORT_COL_PROJECTED_DUE_DATE), _
             ws.Cells(IMPORT_LAST_DATA_ROW, IMPORT_COL_PROJECTED_DUE_DATE)).NumberFormat = "DD/MM/YYYY"
    ws.Range("M1").Value = "Last Import"
    ws.Range("M1").Font.Bold = True
    ws.Range("M2").NumberFormat = "DD/MM/YYYY HH:MM"
    ws.Columns("M").ColumnWidth = 20

    On Error Resume Next
    ws.Activate
    ws.Range("B2").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0

    Set CreateImportSheet = ws
End Function

'------------------------------------------------------------------------------
' Purpose : Creates and formats a new aircraft chart worksheet.
' Input : tail - aircraft tail number.
' Output : Newly created chart worksheet.
'------------------------------------------------------------------------------
Public Function CreateChartSheet(tail As String) As Worksheet
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))

    ws.Name = tail & CHART_SHEET_SUFFIX

    FormatChartSheet ws
    AddChartButtons ws

    Set CreateChartSheet = ws
End Function

'------------------------------------------------------------------------------
' Purpose : Applies standard formatting to an aircraft chart worksheet.
' Input : ws - chart worksheet to format.
' Output : Updates sheet title, colours, column widths, row heights, and freeze panes.
'------------------------------------------------------------------------------
Private Sub FormatChartSheet(ws As Worksheet)
    
    Dim tail As String
    tail = Replace(ws.Name, CHART_SHEET_SUFFIX, "")
    
    ws.Tab.Color = RGB(47, 84, 150)

    With ws.Range("A1")
        .Value = tail & " Maintenance Forecast"
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(47, 84, 150)
    End With

    ws.Range("A1:C1").Merge

    ws.Range("A2").Interior.Color = RGB(0, 32, 96)
    ws.Range("B2").Interior.Color = RGB(255, 255, 0)
    ws.Range("C2").Interior.Color = RGB(255, 0, 0)

    ws.Columns(1).ColumnWidth = 13
    ws.Columns(2).ColumnWidth = 13
    ws.Columns(3).ColumnWidth = 13
    ws.Rows(2).RowHeight = 48
    
    On Error Resume Next
    ws.Activate
    ws.Range("D5").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    
End Sub

'------------------------------------------------------------------------------
' Purpose : Adds standard dashboard, work pack, and reimport buttons to a chart sheet.
' Input : ws - chart worksheet to receive buttons.
' Output : Deletes old matching buttons and creates new ones.
'------------------------------------------------------------------------------
Private Sub AddChartButtons(ws As Worksheet)

    Dim tail As String
    tail = Replace(ws.Name, CHART_SHEET_SUFFIX, "")

    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name Like tail & "_btn*" Or shp.Name Like "*_btnReturn" Or shp.Name Like "*_btnWorkPack" Or shp.Name Like "*_btnReimport" Then
            shp.Delete
        End If
    Next shp

    Dim btnTop As Long
    Dim btnLeft As Long
    Dim btnWidth As Long
    Dim btnHeight As Long
    Dim btnGap As Long

    btnTop = 19.5
    btnLeft = 5
    btnWidth = 65
    btnHeight = 44
    btnGap = 10

    Dim btnReturn As Button
    Set btnReturn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)

    With btnReturn
        .Caption = "Return to Dashboard"
        .OnAction = "GoToDashboard"
        .Name = tail & "_btnReturn"
        .Font.Size = 9
    End With

    Dim btnWorkPack As Button
    Set btnWorkPack = ws.Buttons.Add(btnLeft + btnWidth + btnGap, btnTop, btnWidth, btnHeight)

    With btnWorkPack
        .Caption = "Generate Work Pack"
        .OnAction = "GenerateWordWorkPack"
        .Name = tail & "_btnWorkPack"
        .Font.Size = 9
    End With
    
    Dim btnReimport As Button
    Set btnReimport = ws.Buttons.Add((btnLeft + btnWidth * 2) + (btnGap * 2), btnTop, btnWidth, btnHeight)

    With btnReimport
        .Caption = "Reimport SITS"
        .OnAction = "ReimportFromChartButton"
        .Name = tail & "_btnReimport"
        .Font.Size = 9
    End With

End Sub

'------------------------------------------------------------------------------
' Purpose : Reapplies chart formatting and buttons to all aircraft chart sheets.
' Input : None.
' Output : Updates every worksheet whose name ends with CHART_SHEET_SUFFIX.
'------------------------------------------------------------------------------
Public Sub ReformatAllChartSheets()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
        If Right$(ws.Name, Len(CHART_SHEET_SUFFIX)) = CHART_SHEET_SUFFIX Then
            FormatChartSheet ws
            AddChartButtons ws
        End If
    Next ws
End Sub

'------------------------------------------------------------------------------
' Purpose : Adds aircraft chart navigation hyperlinks across the top of a chart sheet.
' Input : chartWs - chart worksheet receiving navigation links.
' tailNumbers - array of aircraft tail numbers.
' aircraftCount - number of valid aircraft entries.
' Output : Clears and rebuilds navigation links in E1:AZ1.
'------------------------------------------------------------------------------
Public Sub AddChartNavigation(chartWs As Worksheet, tailNumbers() As String, aircraftCount As Long)

    Dim c As Long
    Dim i As Long
    Dim navCell As Range
    Dim targetSheetName As String
    Dim displayTail As String

    With chartWs.Range("E1:AZ1")
        .ClearContents
        .ClearFormats
    End With

    Dim h As Hyperlink
    For Each h In chartWs.Hyperlinks
        If Not Intersect(h.Range, chartWs.Range("E1:AZ1")) Is Nothing Then
            h.Delete
        End If
    Next h

    With chartWs.Range("E1")
        .Value = "Jump to:"
        .Font.Name = "Arial"
        .Font.Size = 9
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With

    c = 6

    For i = 1 To aircraftCount

        displayTail = Trim(CStr(tailNumbers(i)))

        If Len(displayTail) > 0 Then

            targetSheetName = displayTail & CHART_SHEET_SUFFIX

            If SheetExists(targetSheetName) Then

                Set navCell = chartWs.Cells(1, c)

                chartWs.Hyperlinks.Add _
                    Anchor:=navCell, _
                    Address:="", _
                    SubAddress:="'" & targetSheetName & "'!A1", _
                    TextToDisplay:=displayTail

                With navCell
                    .Font.Name = "Arial"
                    .Font.Size = 9
                    .Font.Bold = False
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With

                c = c + 1

            End If

        End If

    Next i

    chartWs.Rows(1).RowHeight = 18

End Sub

'------------------------------------------------------------------------------
' Purpose : Shows or hides all aircraft import sheets as a group.
' Input : None.
' Output : Toggles all sheets ending with IMPORT_SHEET_SUFFIX between visible and hidden.
'------------------------------------------------------------------------------
Public Sub ToggleImportSheetsVisibility()
    Dim ws As Worksheet
    Dim allVisible As Boolean
    allVisible = True
    
    For Each ws In ThisWorkbook.Worksheets
        If Right$(ws.Name, Len(IMPORT_SHEET_SUFFIX)) = IMPORT_SHEET_SUFFIX Then
            If ws.Visible <> xlSheetVisible Then
                allVisible = False
                Exit For
            End If
        End If
    Next ws
    
    For Each ws In ThisWorkbook.Worksheets
        If Right$(ws.Name, Len(IMPORT_SHEET_SUFFIX)) = IMPORT_SHEET_SUFFIX Then
            If allVisible Then
                ws.Visible = xlSheetHidden
            Else
                ws.Visible = xlSheetVisible
            End If
        End If
    Next ws
End Sub

'------------------------------------------------------------------------------
' Purpose : Activates the Dashboard worksheet.
' Input : None.
' Output : Dashboard sheet is activated if it exists.
'------------------------------------------------------------------------------
Public Sub GoToDashboard()
    On Error Resume Next
    ThisWorkbook.Sheets("Dashboard").Activate
    On Error GoTo 0
End Sub
