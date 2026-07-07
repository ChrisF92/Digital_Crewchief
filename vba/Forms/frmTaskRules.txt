Attribute VB_Name = "frmTaskRules"
Attribute VB_Base = "0{BC4C1179-44E1-4116-AF54-8C1BFA873576}{287EDEFE-B71C-4D17-92EA-4C8C8EF68C7D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'------------------------------------------------------------------------------
' Expected controls (layout in frmTaskRules.frx):
'   lblIntro, optHighlight, optNoExtend, optNoPull  (same GroupName, e.g. RuleSection)
'   lblSectionHelp, lstPrefixes, lblPrefixCount
'   lblAddPrefix, txtPrefix, cmdAdd, cmdRemove, cmdApply, cmdCancel
'------------------------------------------------------------------------------

Private Const RULE_SECTION_HIGHLIGHT As String = "Highlight"
Private Const RULE_SECTION_NO_EXTEND As String = "No Extend"
Private Const RULE_SECTION_NO_PULL As String = "No Pull"

Private Const HELP_HIGHLIGHT As String = _
    "Tasks whose codes start with these prefixes are highlighted on maintenance charts."

Private Const HELP_NO_EXTEND As String = _
    "Tasks matching these prefixes are excluded from extension logic."

Private Const HELP_NO_PULL As String = _
    "Tasks matching these prefixes are excluded from pull-forward grouping."

Private m_HighlightPrefixes As Collection
Private m_NoExtendPrefixes As Collection
Private m_NoPullPrefixes As Collection
Private m_IsLoading As Boolean

Private Sub UserForm_Initialize()

    On Error GoTo ErrHandler

    m_IsLoading = True

    Me.Caption = "Task Rules"

    Set m_HighlightPrefixes = LoadTaskRulePrefixes(TASK_RULE_COL_HIGHLIGHT)
    Set m_NoExtendPrefixes = LoadTaskRulePrefixes(TASK_RULE_COL_NO_EXTEND)
    Set m_NoPullPrefixes = LoadTaskRulePrefixes(TASK_RULE_COL_NO_PULL)

    ConfigureFormCaptions

    optHighlight.Value = True
    optNoExtend.Value = False
    optNoPull.Value = False

    UpdateSectionUI

    m_IsLoading = False

    Exit Sub

ErrHandler:
    m_IsLoading = False

    MsgBox "Task Rules failed to open." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Task Rules"

End Sub

Private Sub optHighlight_Click()
    OnRuleSectionChanged
End Sub

Private Sub optNoExtend_Click()
    OnRuleSectionChanged
End Sub

Private Sub optNoPull_Click()
    OnRuleSectionChanged
End Sub

Private Sub lstPrefixes_Click()
    UpdateRemoveButtonState
End Sub

Private Sub txtPrefix_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

    If KeyCode = vbKeyReturn Then
        cmdAdd_Click
    End If

End Sub

Private Sub cmdAdd_Click()

    Dim normalizedPrefix As String
    normalizedPrefix = NormalizeTaskRulePrefix(txtPrefix.Value)

    If Len(normalizedPrefix) = 0 Then
        MsgBox "Please enter a task code prefix.", vbExclamation, "Task Rules"
        txtPrefix.SetFocus
        Exit Sub
    End If

    Dim activePrefixes As Collection
    Set activePrefixes = GetActivePrefixCollection()

    If PrefixExistsInCollection(activePrefixes, normalizedPrefix) Then
        MsgBox "That prefix is already listed in " & ActiveRuleSectionName() & ".", _
               vbExclamation, _
               "Task Rules"
        txtPrefix.SetFocus
        Exit Sub
    End If

    activePrefixes.Add normalizedPrefix
    UpdateSectionUI

    txtPrefix.Value = vbNullString
    txtPrefix.SetFocus

End Sub

Private Sub cmdRemove_Click()

    If lstPrefixes.listIndex < 0 Then
        MsgBox "Please select a prefix to remove.", vbExclamation, "Task Rules"
        Exit Sub
    End If

    Dim activePrefixes As Collection
    Set activePrefixes = GetActivePrefixCollection()

    If lstPrefixes.listIndex + 1 <= activePrefixes.Count Then
        activePrefixes.Remove lstPrefixes.listIndex + 1
    End If

    UpdateSectionUI

End Sub

Private Sub cmdApply_Click()

    If Not ValidatePrefixCollection(m_HighlightPrefixes, RULE_SECTION_HIGHLIGHT) Then Exit Sub
    If Not ValidatePrefixCollection(m_NoExtendPrefixes, RULE_SECTION_NO_EXTEND) Then Exit Sub
    If Not ValidatePrefixCollection(m_NoPullPrefixes, RULE_SECTION_NO_PULL) Then Exit Sub

    If Not ConfirmTaskRulesApply(m_HighlightPrefixes, m_NoExtendPrefixes, m_NoPullPrefixes) Then
        Exit Sub
    End If

    On Error GoTo SaveErrHandler

    SaveAllTaskRules m_HighlightPrefixes, m_NoExtendPrefixes, m_NoPullPrefixes

    Unload Me

    RefreshAll

    Exit Sub

SaveErrHandler:

    MsgBox "Could not save task rules." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Task Rules"

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub OnRuleSectionChanged()

    If m_IsLoading Then Exit Sub

    UpdateSectionUI

End Sub

Private Sub UpdateSectionUI()

    UpdateSectionHelpText
    LoadActiveSectionIntoListBox
    UpdatePrefixCountLabel
    UpdateRemoveButtonState

End Sub

Private Sub UpdateSectionHelpText()

    Select Case ActiveRuleSectionName()

        Case RULE_SECTION_HIGHLIGHT
            lblSectionHelp.Caption = HELP_HIGHLIGHT

        Case RULE_SECTION_NO_EXTEND
            lblSectionHelp.Caption = HELP_NO_EXTEND

        Case RULE_SECTION_NO_PULL
            lblSectionHelp.Caption = HELP_NO_PULL

        Case Else
            lblSectionHelp.Caption = vbNullString

    End Select

End Sub

Private Sub UpdatePrefixCountLabel()

    Dim prefixCount As Long
    prefixCount = GetActivePrefixCollection().Count

    If prefixCount = 1 Then
        lblPrefixCount.Caption = "1 prefix"
    Else
        lblPrefixCount.Caption = CStr(prefixCount) & " prefixes"
    End If

End Sub

Private Sub UpdateRemoveButtonState()

    cmdRemove.Enabled = (lstPrefixes.listIndex >= 0)

End Sub

Private Sub ConfigureFormCaptions()

    On Error Resume Next

    lblIntro.Caption = "Task codes match by prefix (start of code)."

    cmdApply.Caption = "Save Changes"
    cmdCancel.Caption = "Close"
    cmdRemove.Caption = "Remove selected"
    lblAddPrefix.Caption = "Add prefix:"

    On Error GoTo 0

End Sub

Private Sub LoadActiveSectionIntoListBox()

    Dim activePrefixes As Collection
    Set activePrefixes = GetActivePrefixCollection()

    lstPrefixes.Clear

    Dim prefixIndex As Long

    For prefixIndex = 1 To activePrefixes.Count
        lstPrefixes.AddItem activePrefixes(prefixIndex)
    Next prefixIndex

End Sub

Private Function GetActivePrefixCollection() As Collection

    Select Case ActiveRuleSectionName()

        Case RULE_SECTION_HIGHLIGHT
            Set GetActivePrefixCollection = m_HighlightPrefixes

        Case RULE_SECTION_NO_EXTEND
            Set GetActivePrefixCollection = m_NoExtendPrefixes

        Case RULE_SECTION_NO_PULL
            Set GetActivePrefixCollection = m_NoPullPrefixes

        Case Else
            Set GetActivePrefixCollection = m_HighlightPrefixes

    End Select

End Function

Private Function ActiveRuleSectionName() As String

    If optNoExtend.Value Then
        ActiveRuleSectionName = RULE_SECTION_NO_EXTEND
        Exit Function
    End If

    If optNoPull.Value Then
        ActiveRuleSectionName = RULE_SECTION_NO_PULL
        Exit Function
    End If

    ActiveRuleSectionName = RULE_SECTION_HIGHLIGHT

End Function

Private Function PrefixExistsInCollection(ByVal prefixes As Collection, _
                                          ByVal normalizedPrefix As String) As Boolean

    Dim prefixIndex As Long

    For prefixIndex = 1 To prefixes.Count

        If StrComp(CStr(prefixes(prefixIndex)), normalizedPrefix, vbBinaryCompare) = 0 Then
            PrefixExistsInCollection = True
            Exit Function
        End If

    Next prefixIndex

    PrefixExistsInCollection = False

End Function

Private Function ValidatePrefixCollection(ByVal prefixes As Collection, _
                                          ByVal sectionName As String) As Boolean

    ValidatePrefixCollection = False

    Dim seenPrefixes As Object
    Set seenPrefixes = CreateObject("Scripting.Dictionary")

    Dim prefixIndex As Long
    Dim normalizedPrefix As String

    For prefixIndex = 1 To prefixes.Count

        normalizedPrefix = NormalizeTaskRulePrefix(CStr(prefixes(prefixIndex)))

        If Len(normalizedPrefix) = 0 Then
            MsgBox sectionName & " contains a blank prefix.", _
                   vbExclamation, _
                   "Task Rules"
            Exit Function
        End If

        If seenPrefixes.Exists(normalizedPrefix) Then
            MsgBox "Duplicate prefix """ & normalizedPrefix & """ in " & sectionName & ".", _
                   vbExclamation, _
                   "Task Rules"
            Exit Function
        End If

        seenPrefixes.Add normalizedPrefix, True

    Next prefixIndex

    ValidatePrefixCollection = True

End Function

Private Function ConfirmTaskRulesApply(ByVal highlightPrefixes As Collection, _
                                       ByVal noExtendPrefixes As Collection, _
                                       ByVal noPullPrefixes As Collection) As Boolean

    Dim messageText As String

    messageText = "You are about to save task rules." & vbCrLf & vbCrLf & _
                  RULE_SECTION_HIGHLIGHT & ": " & highlightPrefixes.Count & " prefix(es)" & vbCrLf & _
                  RULE_SECTION_NO_EXTEND & ": " & noExtendPrefixes.Count & " prefix(es)" & vbCrLf & _
                  RULE_SECTION_NO_PULL & ": " & noPullPrefixes.Count & " prefix(es)" & vbCrLf & vbCrLf & _
                  "After you confirm, the rules will be saved, then the charts and dashboard will be refreshed automatically." & vbCrLf & vbCrLf & _
                  "Do you want to continue?"

    ConfirmTaskRulesApply = _
        (MsgBox(messageText, vbQuestion + vbYesNo, "Confirm Task Rules Update") = vbYes)

End Function

