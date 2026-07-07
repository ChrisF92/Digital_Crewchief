Attribute VB_Name = "modApplicationState"
Option Explicit

'------------------------------------------------------------------------------
' Module : modApplicationState
' Purpose : Provides safe helpers for temporarily changing Excel application
' settings during long-running procedures.
'------------------------------------------------------------------------------

Private mPreviousScreenUpdating As Boolean
Private mPreviousEnableEvents As Boolean
Private mPreviousDisplayAlerts As Boolean
Private mPreviousCalculation As XlCalculation

Private mStateCaptured As Boolean


'------------------------------------------------------------------------------
' Purpose : Captures the current Excel application state and switches Excel into
' a faster mode for macro execution.
' Notes : Call EndFastMode in the procedure's CleanExit and ErrHandler paths.
'------------------------------------------------------------------------------
Public Sub BeginFastMode(Optional ByVal suppressAlerts As Boolean = False)

    If Not mStateCaptured Then
        mPreviousScreenUpdating = Application.ScreenUpdating
        mPreviousEnableEvents = Application.EnableEvents
        mPreviousDisplayAlerts = Application.DisplayAlerts
        mPreviousCalculation = Application.Calculation

        mStateCaptured = True
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If suppressAlerts Then
        Application.DisplayAlerts = False
    End If

End Sub


'------------------------------------------------------------------------------
' Purpose : Restores Excel application settings captured by BeginFastMode.
'------------------------------------------------------------------------------
Public Sub EndFastMode()

    If Not mStateCaptured Then Exit Sub

    Application.ScreenUpdating = mPreviousScreenUpdating
    Application.EnableEvents = mPreviousEnableEvents
    Application.DisplayAlerts = mPreviousDisplayAlerts
    Application.Calculation = mPreviousCalculation

    mStateCaptured = False

End Sub


'------------------------------------------------------------------------------
' Purpose : Forces Excel back to a normal interactive state.
' Notes : Useful as a manual recovery macro if a procedure is interrupted.
'------------------------------------------------------------------------------
Public Sub ResetApplicationState()

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic

    mStateCaptured = False

End Sub

