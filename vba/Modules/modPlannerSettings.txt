Attribute VB_Name = "modPlannerSettings"
Option Explicit

'------------------------------------------------------------------------------
' Module : modPlannerSettings
' Purpose : Loads and validates user-configurable planner settings.
'------------------------------------------------------------------------------

Private Const ERR_BASE_PLANNER_SETTINGS As Long = vbObjectError + 2100

'------------------------------------------------------------------------------
' Type : PlannerSettings
' Purpose : Holds all user-configurable settings needed during a refresh.
'------------------------------------------------------------------------------
Public Type PlannerSettings

    forecastStart As Date
    ForecastEnd As Date
    WeeksShown As Long

    displayStart As Date
    displayEnd As Date
    displayWeeksShown As Long

    MaxAircraftSlots As Long

    flyingWeeks As Long
    MaintenanceWeeks As Long
    cycleLength As Long

    reforecastThreshold As Double
    GroupingTolerancePercent As Double
    MaxExtensionPercent As Double
    MaxPullForwardPercent As Double
    
    WeeksBeforeForecastStart As Long

    DashboardSortOrder As String

End Type


'------------------------------------------------------------------------------
' Purpose : Loads all planner settings from workbook-level named ranges.
' Output : Populated PlannerSettings value.
' Raises : Error if required settings are missing or invalid.
'------------------------------------------------------------------------------
Public Function LoadPlannerSettings() As PlannerSettings

    Dim settings As PlannerSettings

    settings.forecastStart = GetNamedDate("ForecastStart")
    settings.WeeksShown = GetNamedLong("WeeksShown")
    settings.ForecastEnd = settings.forecastStart + 7 * settings.WeeksShown - 1
    settings.WeeksBeforeForecastStart = GetNamedLong("WeeksBeforeForecastStart")

    settings.MaxAircraftSlots = GetNamedLong("MaxAircraftSlots")

    settings.flyingWeeks = GetNamedLong("FlyingWeeks")
    settings.MaintenanceWeeks = GetNamedLong("MaintenanceWeeks")
    settings.cycleLength = settings.flyingWeeks + settings.MaintenanceWeeks

    settings.reforecastThreshold = GetNamedDouble("ReforecastThreshold")
    settings.GroupingTolerancePercent = GetNamedDouble("AlignmentThreshold") / 100
    settings.MaxExtensionPercent = GetNamedDouble("MaxExtensionPercent") / 100
    settings.MaxPullForwardPercent = GetNamedDouble("PullThreshold") / 100

    settings.DashboardSortOrder = LoadDashboardSortOrderSetting()

    ValidatePlannerSettings settings

    settings.displayStart = settings.forecastStart - 7 * settings.WeeksBeforeForecastStart
    settings.displayEnd = settings.ForecastEnd
    settings.displayWeeksShown = settings.WeeksShown + settings.WeeksBeforeForecastStart

    LoadPlannerSettings = settings

End Function


'------------------------------------------------------------------------------
' Purpose : Validates settings after they have been loaded.
' Input : settings - PlannerSettings value to validate.
' Raises : Error with a clear message if any setting is invalid.
'------------------------------------------------------------------------------
Private Sub ValidatePlannerSettings(ByRef settings As PlannerSettings)

    If settings.WeeksShown <= 0 Then
        Err.Raise ERR_BASE_PLANNER_SETTINGS + 1, _
                  "ValidatePlannerSettings", _
                  "WeeksShown must be greater than zero."
    End If

    If settings.MaxAircraftSlots <= 0 Then
        Err.Raise ERR_BASE_PLANNER_SETTINGS + 2, _
                  "ValidatePlannerSettings", _
                  "MaxAircraftSlots must be greater than zero."
    End If

    If settings.flyingWeeks < 0 Then
        Err.Raise ERR_BASE_PLANNER_SETTINGS + 3, _
                  "ValidatePlannerSettings", _
                  "FlyingWeeks cannot be negative."
    End If

    If settings.MaintenanceWeeks < 0 Then
        Err.Raise ERR_BASE_PLANNER_SETTINGS + 4, _
                  "ValidatePlannerSettings", _
                  "MaintenanceWeeks cannot be negative."
    End If

    If settings.cycleLength <= 0 Then
        Err.Raise ERR_BASE_PLANNER_SETTINGS + 5, _
                  "ValidatePlannerSettings", _
                  "FlyingWeeks plus MaintenanceWeeks must be greater than zero."
    End If

    If settings.WeeksBeforeForecastStart < 0 Then
        Err.Raise ERR_BASE_PLANNER_SETTINGS + 6, _
                  "ValidatePlannerSettings", _
                  "WeeksBeforeForecastStart cannot be negative."
    End If

    If settings.GroupingTolerancePercent < 0 Then
        Err.Raise ERR_BASE_PLANNER_SETTINGS + 7, _
                  "ValidatePlannerSettings", _
                  "GroupingTolerancePercent cannot be negative."
    End If

    If settings.MaxExtensionPercent < 0 Then
        Err.Raise ERR_BASE_PLANNER_SETTINGS + 8, _
                  "ValidatePlannerSettings", _
                  "MaxExtensionPercent cannot be negative."
    End If

    If settings.MaxPullForwardPercent < 0 Then
        Err.Raise ERR_BASE_PLANNER_SETTINGS + 9, _
                  "ValidatePlannerSettings", _
                  "MaxPullForwardPercent cannot be negative."
    End If

End Sub


'------------------------------------------------------------------------------
' Purpose : Reads the dashboard aircraft sort setting, defaulting to tail number.
' Output : DASHBOARD_SORT_TAIL or DASHBOARD_SORT_CYCLE.
'------------------------------------------------------------------------------
Private Function LoadDashboardSortOrderSetting() As String

    Dim sortOrder As String

    On Error Resume Next
    sortOrder = Trim$(CStr(GetNamedRangeValue("DashboardSortOrder")))
    On Error GoTo 0

    If StrComp(sortOrder, DASHBOARD_SORT_CYCLE, vbTextCompare) = 0 Then
        LoadDashboardSortOrderSetting = DASHBOARD_SORT_CYCLE
    Else
        LoadDashboardSortOrderSetting = DASHBOARD_SORT_TAIL
    End If

End Function


