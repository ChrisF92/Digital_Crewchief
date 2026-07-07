Attribute VB_Name = "ExportAllComponents"
Option Explicit

'------------------------------------------------------------------------------
' Export all VBA modules and UserForms from this workbook into the vba/ folder.
'
' Prerequisites (Excel Trust Center):
'   File -> Options -> Trust Center -> Trust Center Settings ->
'   Macro Settings -> Trust access to the VBA project object model
'
' Usage:
'   1. Open 20260603-Digital Crewchief_BETA_MASTER.xlsm
'   2. Alt+F11 -> File -> Import File -> select this module
'   3. Alt+F8 -> Run ExportAllComponentsToRepo
'------------------------------------------------------------------------------

Public Sub ExportAllComponentsToRepo()
    Dim repoRoot As String
    repoRoot = ThisWorkbook.Path

    If Len(repoRoot) = 0 Then
        MsgBox "Save the workbook before exporting.", vbExclamation, "Export VBA"
        Exit Sub
    End If

    EnsureFolder repoRoot & "\vba\Modules"
    EnsureFolder repoRoot & "\vba\ExcelObjects"
    EnsureFolder repoRoot & "\vba\Forms"

    Dim comp As Object
    Dim exportPath As String
    Dim frxPath As String
    Dim exportedCount As Long

    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1 ' vbext_ct_StdModule
                exportPath = repoRoot & "\vba\Modules\" & comp.Name & ".bas"
            Case 2 ' vbext_ct_ClassModule
                If Left$(comp.Name, 3) = "frm" Then
                    exportPath = repoRoot & "\vba\Forms\" & comp.Name & ".frm"
                Else
                    exportPath = repoRoot & "\vba\ExcelObjects\" & comp.Name & ".cls"
                End If
            Case 3 ' vbext_ct_MSForm
                exportPath = repoRoot & "\vba\Forms\" & comp.Name & ".frm"
            Case Else
                GoTo NextComponent
        End Select

        frxPath = Left$(exportPath, Len(exportPath) - 4) & ".frx"
        On Error Resume Next
        Kill exportPath
        Kill frxPath
        On Error GoTo 0

        comp.Export exportPath
        exportedCount = exportedCount + 1

NextComponent:
    Next comp

    MsgBox "Exported " & exportedCount & " VBA components to:" & vbCrLf & _
           repoRoot & "\vba", vbInformation, "Export VBA"
End Sub

Private Sub EnsureFolder(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub
