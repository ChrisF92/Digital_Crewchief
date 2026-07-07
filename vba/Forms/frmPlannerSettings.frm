Attribute VB_Name = "frmPlannerSettings"
Attribute VB_Base = "0{4FA4E134-221A-4384-8C47-C49239004113}{E8E87C65-2195-4DBE-9719-73A4BB5E2569}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Private Const SETTINGS_SHEET_NAME As String = "Settings"

Private Const FORECAST_MODE_AUTO As String = "Current Week"
Private Const FORECAST_MODE_MANUAL As String = "Manual Date"

Private m_IsLoading As Boolean

' Layout is designed in the VBA UserForm designer.
' Code only toggles control visibility and enabled state.

Private Sub cmbEnableTaskAlignment_Change()

    ApplyPlannerSettingsVisibility

End Sub

Private Sub cmdApply_Click()

    If Not ValidateSettingsForm() Then
        Exit Sub
    End If

    If Not ConfirmSettingsApply() Then
        Exit Sub
    End If

    SaveSettingsFromForm

    Unload Me

    RefreshAll

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub txtManualForecastStart_Change()

    If m_IsLoading Then Exit Sub

    UpdateWeeksBeforeForecastFromManualDate

End Sub

Private Sub UserForm_Initialize()

    On Error GoTo ErrHandler

    m_IsLoading = True

    Me.Caption = "Planner Settings"

    LoadForecastModeList
    LoadYesNoList cmbEnableTaskAlignment
    LoadExtensionThresholdList
    LoadDashboardSortOrderList
    LoadSettingsIntoForm

    m_IsLoading = False

    ApplyPlannerSettingsVisibility

    Exit Sub

ErrHandler:
    m_IsLoading = False

    MsgBox "Planner Settings failed to open." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Planner Settings"

End Sub


Private Sub LoadForecastModeList()

    cmbForecastStartMode.Clear
    cmbForecastStartMode.AddItem FORECAST_MODE_AUTO
    cmbForecastStartMode.AddItem FORECAST_MODE_MANUAL

End Sub

Private Sub LoadExtensionThresholdList()

    cmbExtensionThreshold.Clear
    cmbExtensionThreshold.AddItem "10%"
    cmbExtensionThreshold.AddItem "20%"

End Sub

Private Sub LoadDashboardSortOrderList()

    cmbDashboardSortOrder.Clear
    cmbDashboardSortOrder.AddItem DASHBOARD_SORT_TAIL
    cmbDashboardSortOrder.AddItem DASHBOARD_SORT_CYCLE

End Sub

Private Sub LoadSettingsIntoForm()

    ' Forecast settings
    cmbForecastStartMode.Value = CStr(GetNamedRangeValue("ForecastStartMode"))
    
    If StrComp(cmbForecastStartMode.Value, FORECAST_MODE_MANUAL, vbTextCompare) = 0 Then
        txtManualForecastStart.Value = FormatDateForForm(GetNamedRangeValue("ForecastStart"))
    Else
        txtManualForecastStart.Value = vbNullString
    End If
    
    txtForecastLength.Value = CStr(GetNamedRangeValue("WeeksShown"))
    txtWeeksBeforeForecast.Value = CStr(GetNamedRangeValue("WeeksBeforeForecastStart"))
    
    'Task settings
    cmbEnableTaskAlignment.Value = BooleanToYesNo(GetNamedRangeValue("EnableTaskAlignment"))
    txtAlignmentThreshold.Value = CStr(GetNamedRangeValue("AlignmentThreshold"))
    txtPullThreshold.Value = CStr(GetNamedRangeValue("PullThreshold"))
    cmbExtensionThreshold.Value = FormatPercentOption(GetNamedRangeValue("MaxExtensionPercent"))

    cmbDashboardSortOrder.Value = LoadDashboardSortOrderFormValue()

    ' Fixed cycle structure
    txtCycleLength.Value = CStr(GetNamedRangeValue("CycleLength"))
    txtFlyingWeeks.Value = CStr(GetNamedRangeValue("FlyingWeeks"))
    txtMaintenanceWeeks.Value = CStr(GetNamedRangeValue("MaintenanceWeeks"))

    txtCycleLength.Locked = True
    txtFlyingWeeks.Locked = True
    txtMaintenanceWeeks.Locked = True

    ' Flying rates as HH:MM
    LoadHoursMinutesSetting txtFW1HH, "FlyingWeek1HHRate"
    LoadHoursMinutesSetting txtFW1E1, "FlyingWeek1E1Rate"
    LoadHoursMinutesSetting txtFW1E2, "FlyingWeek1E2Rate"

    LoadHoursMinutesSetting txtFW2HH, "FlyingWeek2HHRate"
    LoadHoursMinutesSetting txtFW2E1, "FlyingWeek2E1Rate"
    LoadHoursMinutesSetting txtFW2E2, "FlyingWeek2E2Rate"

    LoadHoursMinutesSetting txtFW3HH, "FlyingWeek3HHRate"
    LoadHoursMinutesSetting txtFW3E1, "FlyingWeek3E1Rate"
    LoadHoursMinutesSetting txtFW3E2, "FlyingWeek3E2Rate"

End Sub

Private Sub cmbForecastStartMode_Change()

    ApplyPlannerSettingsVisibility

End Sub

Private Sub UpdateForecastFieldAvailability()

    Dim isManual As Boolean
    isManual = (StrComp(Trim$(cmbForecastStartMode.Value), FORECAST_MODE_MANUAL, vbTextCompare) = 0)

    lblManualForecastStart.Visible = isManual
    txtManualForecastStart.Visible = isManual

    lblWeeksBeforeForecast.Visible = isManual
    txtWeeksBeforeForecast.Visible = isManual

    If Not isManual Then
        If Not m_IsLoading Then
            txtManualForecastStart.Value = vbNullString
            txtWeeksBeforeForecast.Value = vbNullString
        End If
    Else
        If Not m_IsLoading Then
            UpdateWeeksBeforeForecastFromManualDate
        End If
    End If

End Sub


Private Sub UpdateTaskGroupingFieldAvailability()

    Dim isEnabled As Boolean
    isEnabled = YesNoToBoolean(cmbEnableTaskAlignment.Value)

    lblAlignmentThreshold.Visible = isEnabled
    txtAlignmentThreshold.Visible = isEnabled

    If Not isEnabled Then
        If Not m_IsLoading Then
            txtAlignmentThreshold.Value = vbNullString
        End If
    End If

End Sub

Private Function ValidateYesNoComboBox(ByVal comboBoxControl As Object, _
                                       ByVal displayName As String) As Boolean

    ValidateYesNoComboBox = False

    If StrComp(comboBoxControl.Value, "Yes", vbTextCompare) <> 0 And _
       StrComp(comboBoxControl.Value, "No", vbTextCompare) <> 0 Then

        MsgBox displayName & " must be Yes or No.", _
               vbExclamation, _
               "Planner Settings"

        comboBoxControl.SetFocus
        Exit Function
    End If

    ValidateYesNoComboBox = True

End Function

Private Function ValidateExtensionThresholdComboBox() As Boolean

    ValidateExtensionThresholdComboBox = False
    
    If cmbExtensionThreshold.Value <> "10%" And cmbExtensionThreshold.Value <> "20%" Then
        MsgBox "Extension Threshold must be 10% or 20%.", _
            vbExclamation, _
            "Planner Settings"
        cmbExtensionThreshold.SetFocus
        Exit Function
    End If
                
    ValidateExtensionThresholdComboBox = True
    
End Function

Private Function ValidateDashboardSortOrderComboBox() As Boolean

    ValidateDashboardSortOrderComboBox = False

    If StrComp(cmbDashboardSortOrder.Value, DASHBOARD_SORT_TAIL, vbTextCompare) <> 0 And _
       StrComp(cmbDashboardSortOrder.Value, DASHBOARD_SORT_CYCLE, vbTextCompare) <> 0 Then

        MsgBox "Dashboard Order must be Tail Number or Cycle Week.", _
               vbExclamation, _
               "Planner Settings"
        cmbDashboardSortOrder.SetFocus
        Exit Function
    End If

    ValidateDashboardSortOrderComboBox = True

End Function

Private Function LoadDashboardSortOrderFormValue() As String

    Dim sortOrder As String

    On Error Resume Next
    sortOrder = Trim$(CStr(GetNamedRangeValue("DashboardSortOrder")))
    On Error GoTo 0

    If StrComp(sortOrder, DASHBOARD_SORT_CYCLE, vbTextCompare) = 0 Then
        LoadDashboardSortOrderFormValue = DASHBOARD_SORT_CYCLE
    Else
        LoadDashboardSortOrderFormValue = DASHBOARD_SORT_TAIL
    End If

End Function

Private Sub UpdateWeeksBeforeForecastFromManualDate()

    If StrComp(Trim$(cmbForecastStartMode.Value), FORECAST_MODE_MANUAL, vbTextCompare) <> 0 Then
        txtWeeksBeforeForecast.Value = vbNullString
        Exit Sub
    End If

    Dim manualForecastStart As Date

    If Not TryGetDate(txtManualForecastStart.Value, manualForecastStart) Then
        txtWeeksBeforeForecast.Value = vbNullString
        Exit Sub
    End If

    If Weekday(manualForecastStart, vbMonday) <> 1 Then
        txtWeeksBeforeForecast.Value = vbNullString
        Exit Sub
    End If

    Dim currentWeekMonday As Date
    currentWeekMonday = Date - Weekday(Date, vbMonday) + 1

    Dim weekDifference As Long
    weekDifference = DateDiff("ww", currentWeekMonday, manualForecastStart, vbMonday, vbFirstFourDays)

    If weekDifference < 0 Then
        txtWeeksBeforeForecast.Value = vbNullString
    Else
        txtWeeksBeforeForecast.Value = CStr(weekDifference)
    End If

End Sub

Private Function ValidateSettingsForm() As Boolean

    ValidateSettingsForm = False

    If Len(Trim$(cmbForecastStartMode.Value)) = 0 Then
        MsgBox "Please select a Forecast Start Mode.", vbExclamation, "Planner Settings"
        cmbForecastStartMode.SetFocus
        Exit Function
    End If

    If StrComp(cmbForecastStartMode.Value, FORECAST_MODE_MANUAL, vbTextCompare) = 0 Then

        Dim manualForecastStart As Date

        If Not TryGetDate(txtManualForecastStart.Value, manualForecastStart) Then
            MsgBox "Please enter a valid Manual Forecast Start date.", vbExclamation, "Planner Settings"
            txtManualForecastStart.SetFocus
            Exit Function
        End If

        If Weekday(manualForecastStart, vbMonday) <> 1 Then
            MsgBox "Manual Forecast Start must be a Monday.", vbExclamation, "Planner Settings"
            txtManualForecastStart.SetFocus
            Exit Function
        End If
        
        Dim currentWeekMonday As Date
        currentWeekMonday = Date - Weekday(Date, vbMonday) + 1
        
        If manualForecastStart < currentWeekMonday Then
            MsgBox "Manual Forecast Start cannot be before the current week Monday.", _
                   vbExclamation, _
                   "Planner Settings"
            txtManualForecastStart.SetFocus
            Exit Function
        End If

    End If

    If Not ValidateWholeNumberTextBox(txtForecastLength, "Forecast Length", 1, 520) Then Exit Function
    
    If StrComp(cmbForecastStartMode.Value, FORECAST_MODE_MANUAL, vbTextCompare) = 0 Then
        If Not ValidateWholeNumberTextBox(txtWeeksBeforeForecast, "Weeks Before Forecast Start", 0, 520) Then Exit Function
    End If

    If Not ValidateHoursMinutesTextBox(txtFW1HH, "Flying Week 1 HH") Then Exit Function
    If Not ValidateHoursMinutesTextBox(txtFW1E1, "Flying Week 1 E1") Then Exit Function
    If Not ValidateHoursMinutesTextBox(txtFW1E2, "Flying Week 1 E2") Then Exit Function

    If Not ValidateHoursMinutesTextBox(txtFW2HH, "Flying Week 2 HH") Then Exit Function
    If Not ValidateHoursMinutesTextBox(txtFW2E1, "Flying Week 2 E1") Then Exit Function
    If Not ValidateHoursMinutesTextBox(txtFW2E2, "Flying Week 2 E2") Then Exit Function

    If Not ValidateHoursMinutesTextBox(txtFW3HH, "Flying Week 3 HH") Then Exit Function
    If Not ValidateHoursMinutesTextBox(txtFW3E1, "Flying Week 3 E1") Then Exit Function
    If Not ValidateHoursMinutesTextBox(txtFW3E2, "Flying Week 3 E2") Then Exit Function
    
    If Not ValidateYesNoComboBox(cmbEnableTaskAlignment, "Enable Task Alignment") Then Exit Function

    If YesNoToBoolean(cmbEnableTaskAlignment.Value) Then
        If Not ValidateWholeNumberTextBox(txtAlignmentThreshold, "Alignment Threshold", 0, 52) Then Exit Function
    End If

    If Not ValidateWholeNumberTextBox(txtPullThreshold, "Pull Threshold", 0, 52) Then Exit Function
    
    If Not ValidateExtensionThresholdComboBox() Then Exit Function

    If Not ValidateDashboardSortOrderComboBox() Then Exit Function

    ValidateSettingsForm = True

End Function

Private Sub SaveSettingsFromForm()

    SetNamedRangeValue "ForecastStartMode", cmbForecastStartMode.Value

    If StrComp(cmbForecastStartMode.Value, FORECAST_MODE_MANUAL, vbTextCompare) = 0 Then
        SetNamedRangeValue "ForecastStart", CDate(txtManualForecastStart.Value)
        SetNamedRangeValue "WeeksBeforeForecastStart", CLng(txtWeeksBeforeForecast.Value)
    Else
        SetNamedRangeValue "ForecastStart", GetCurrentWeekMonday()
        SetNamedRangeValue "WeeksBeforeForecastStart", 0
    End If

    SetNamedRangeValue "WeeksShown", CLng(txtForecastLength.Value)

    SaveHoursMinutesSetting txtFW1HH, "FlyingWeek1HHRate"
    SaveHoursMinutesSetting txtFW1E1, "FlyingWeek1E1Rate"
    SaveHoursMinutesSetting txtFW1E2, "FlyingWeek1E2Rate"

    SaveHoursMinutesSetting txtFW2HH, "FlyingWeek2HHRate"
    SaveHoursMinutesSetting txtFW2E1, "FlyingWeek2E1Rate"
    SaveHoursMinutesSetting txtFW2E2, "FlyingWeek2E2Rate"

    SaveHoursMinutesSetting txtFW3HH, "FlyingWeek3HHRate"
    SaveHoursMinutesSetting txtFW3E1, "FlyingWeek3E1Rate"
    SaveHoursMinutesSetting txtFW3E2, "FlyingWeek3E2Rate"
    SetNamedRangeValue "EnableTaskAlignment", YesNoToBoolean(cmbEnableTaskAlignment.Value)

    If YesNoToBoolean(cmbEnableTaskAlignment.Value) Then
        SetNamedRangeValue "AlignmentThreshold", CLng(txtAlignmentThreshold.Value)
        SetNamedRangeValue "PullThreshold", CLng(txtPullThreshold.Value)
    Else
        SetNamedRangeValue "AlignmentThreshold", 0
        SetNamedRangeValue "PullThreshold", 0
    End If
    
    SetNamedRangeValue "MaxExtensionPercent", PercentTextToNumber(cmbExtensionThreshold.Value)

    SetNamedRangeValue "DashboardSortOrder", cmbDashboardSortOrder.Value

End Sub

Private Function ConfirmSettingsApply() As Boolean

    Dim messageText As String

    messageText = "You are about to save planner settings." & vbCrLf & vbCrLf & _
                  "After you confirm, the settings will be saved, then the charts and dashboard will be refreshed automatically." & vbCrLf & vbCrLf & _
                  "Do you want to continue?"

    ConfirmSettingsApply = _
        (MsgBox(messageText, vbQuestion + vbYesNo, "Confirm Settings Update") = vbYes)

End Function

Private Function ValidateHoursMinutesTextBox(ByVal textBoxControl As Object, _
                                             ByVal displayName As String) As Boolean

    ValidateHoursMinutesTextBox = False

    Dim parsedValue As Double

    If Not TryParseHoursMinutes(textBoxControl.Value, parsedValue) Then
        MsgBox displayName & " must be entered as HH:MM." & vbCrLf & vbCrLf & _
               "Example: 12:30 or 125:15", _
               vbExclamation, _
               "Planner Settings"
        textBoxControl.SetFocus
        Exit Function
    End If

    ValidateHoursMinutesTextBox = True

End Function

Private Sub LoadHoursMinutesSetting(ByVal textBoxControl As Object, _
                                    ByVal rangeName As String)

    textBoxControl.Value = FormatHoursMinutes(GetNamedRangeValue(rangeName))

End Sub

Private Sub SaveHoursMinutesSetting(ByVal textBoxControl As Object, _
                                    ByVal rangeName As String)

    Dim rateValue As Double

    If Not TryParseHoursMinutes(textBoxControl.Value, rateValue) Then
        Exit Sub
    End If

    With ThisWorkbook.Names(rangeName).RefersToRange
        .Value = rateValue
        .NumberFormat = "[h]:mm"
    End With

End Sub

Private Function TryParseHoursMinutes(ByVal inputValue As String, _
                                      ByRef outputExcelValue As Double) As Boolean

    Dim cleanedValue As String
    cleanedValue = Trim$(inputValue)

    If Len(cleanedValue) = 0 Then
        TryParseHoursMinutes = False
        Exit Function
    End If

    Dim parts() As String
    parts = Split(cleanedValue, ":")

    If UBound(parts) <> 1 Then
        TryParseHoursMinutes = False
        Exit Function
    End If

    If Not IsNumeric(parts(0)) Or Not IsNumeric(parts(1)) Then
        TryParseHoursMinutes = False
        Exit Function
    End If

    Dim hoursPart As Long
    Dim minutesPart As Long

    hoursPart = CLng(parts(0))
    minutesPart = CLng(parts(1))

    If hoursPart < 0 Then
        TryParseHoursMinutes = False
        Exit Function
    End If

    If minutesPart < 0 Or minutesPart > 59 Then
        TryParseHoursMinutes = False
        Exit Function
    End If

    outputExcelValue = (hoursPart / 24#) + (minutesPart / 1440#)

    TryParseHoursMinutes = True

End Function

Private Function FormatHoursMinutes(ByVal excelValue As Variant) As String

    If Not IsNumeric(excelValue) Then
        FormatHoursMinutes = vbNullString
        Exit Function
    End If

    Dim totalMinutes As Long
    totalMinutes = CLng(Round(CDbl(excelValue) * 1440#, 0))

    Dim hoursPart As Long
    Dim minutesPart As Long

    hoursPart = totalMinutes \ 60
    minutesPart = totalMinutes Mod 60

    FormatHoursMinutes = CStr(hoursPart) & ":" & Format$(minutesPart, "00")

End Function

Private Function ValidateWholeNumberTextBox(ByVal textBoxControl As Object, _
                                            ByVal displayName As String, _
                                            ByVal minimumValue As Long, _
                                            ByVal maximumValue As Long) As Boolean

    ValidateWholeNumberTextBox = False

    If Len(Trim$(textBoxControl.Value)) = 0 Then
        MsgBox displayName & " cannot be blank.", vbExclamation, "Planner Settings"
        textBoxControl.SetFocus
        Exit Function
    End If

    If Not IsNumeric(textBoxControl.Value) Then
        MsgBox displayName & " must be a number.", vbExclamation, "Planner Settings"
        textBoxControl.SetFocus
        Exit Function
    End If

    If CLng(textBoxControl.Value) < minimumValue Or CLng(textBoxControl.Value) > maximumValue Then
        MsgBox displayName & " must be between " & minimumValue & " and " & maximumValue & ".", _
               vbExclamation, _
               "Planner Settings"
        textBoxControl.SetFocus
        Exit Function
    End If

    ValidateWholeNumberTextBox = True

End Function

Private Function FormatDateForForm(ByVal inputValue As Variant) As String

    If IsDate(inputValue) Then
        FormatDateForForm = Format$(CDate(inputValue), "dd/mm/yyyy")
    Else
        FormatDateForForm = vbNullString
    End If

End Function

Private Function GetNamedRangeValue(ByVal rangeName As String) As Variant

    GetNamedRangeValue = ThisWorkbook.Names(rangeName).RefersToRange.Value

End Function

Private Sub SetNamedRangeValue(ByVal rangeName As String, ByVal newValue As Variant)

    On Error GoTo ErrHandler

    Dim targetRange As Range
    Set targetRange = ThisWorkbook.Names(rangeName).RefersToRange

    If targetRange.Cells.CountLarge <> 1 Then
        Err.Raise vbObjectError + 2200, _
                  "SetNamedRangeValue", _
                  "Named range '" & rangeName & "' must refer to one cell only. It currently refers to " & targetRange.Address(External:=True)
    End If

    If targetRange.Worksheet.ProtectContents Then
        Err.Raise vbObjectError + 2201, _
                  "SetNamedRangeValue", _
                  "Cannot write to named range '" & rangeName & "' because sheet '" & targetRange.Worksheet.Name & "' is protected."
    End If

    targetRange.Value = newValue

    Exit Sub

ErrHandler:
    MsgBox "Could not write setting: " & rangeName & vbCrLf & vbCrLf & _
           "Value attempted: " & CStr(newValue) & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Planner Settings"

    Err.Raise Err.Number

End Sub


Private Sub ApplyPlannerSettingsVisibility()

    UpdateForecastFieldAvailability
    UpdateTaskGroupingFieldAvailability

    On Error Resume Next
    lblAlignmentPercentSuffix.Visible = txtAlignmentThreshold.Visible
    lblPullPercentSuffix.Visible = True
    On Error GoTo 0

End Sub


Private Sub LoadYesNoList(ByVal comboBoxControl As Object)

    comboBoxControl.Clear
    comboBoxControl.AddItem "Yes"
    comboBoxControl.AddItem "No"

End Sub

Private Function BooleanToYesNo(ByVal inputValue As Variant) As String

    If CBool(inputValue) Then
        BooleanToYesNo = "Yes"
    Else
        BooleanToYesNo = "No"
    End If

End Function

Private Function YesNoToBoolean(ByVal inputValue As String) As Boolean

    YesNoToBoolean = (StrComp(Trim$(inputValue), "Yes", vbTextCompare) = 0)

End Function



