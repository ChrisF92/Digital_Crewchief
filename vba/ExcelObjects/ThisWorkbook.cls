Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Open()

    ResetChartImportWarningSession

    If HasImportedAircraftData() Then
        ' Defer refresh so Excel finishes loading before charts/dashboard are rebuilt.
        Application.OnTime Now + TimeValue("00:00:01"), "DeferredRefreshAll"
    End If

End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)

    On Error Resume Next

    If TypeName(Sh) = "Worksheet" Then
        HandleChartSheetActivate Sh
    End If

    On Error GoTo 0

End Sub

