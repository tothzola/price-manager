VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MultiFrameViewExtended"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("AppObjects.View")
Option Explicit
'By Kamal Bharakhda

'-----------------------------------------------------------------------
'Settings
'-----------------------------------------------------------------------

Private Const CONTROL_TYPE_FRAME As String = "Frame"
Private Const CONTROL_TYPE_TEXTBOX As String = "TextBox"
Private Const CONTROL_TYPE_LABEL As String = "Label"
Private Const CONTROL_TYPE_COMBOBOX As String = "ComboBox"
Private Const CONTROL_TYPE_LISTBOX As String = "ListBox"
Private Const PANEL_TYPE_ALWAYSON As String = "INFO"
Private Const PANEL_TYPE_SIDE As String = "SIDE"
Private Const PANEL_TYPE_MAIN As String = "MAIN"
Private Const SYMBOL_OK As String = "ü"
Private Const CONTROL_NAME_OK As String = "OK"

'Colors
Private Const INVALID_FIELD_BORDERCOLOR As Long = COLOR_OF_NOT_OKAY
Private Const INVALID_FIELD_FONTCOLOR As Long = COLOR_OF_NOT_OKAY
Private Const VALID_FIELD_BORDERCOLOR As Long = &HC0C0C0
Private Const VALID_FIELD_FONTCOLOR As Long = vbBlack

'-----------------------------------------------------------------------
'Public Properties
'-----------------------------------------------------------------------

'Userform Properties
Public TargetForm As Object 'Instance of User Form/View
Public formWidth As Integer
Public formHeight As Integer

'Properties of a frame that always on
Public frameAlwaysOn As MSForms.Frame
Public alwaysOnTop As Integer
Public alwaysOnLeft As Integer
Public alwaysOnWidth As Integer
Public alwaysOnHeight As Integer
Private IsAlwaysOnFrameVisible As Boolean

'Properties of Side Frame
Public sideFrameTop As Integer
Public sideFrameLeft As Integer
Public sideFrameWidth As Integer
Public sideFrameHeight As Integer

'Properties of Main Frame
Public mainFrameTop As Integer
Public mainFrameLeft As Integer
Public mainFrameWidth As Integer
Public mainFrameHeight As Integer

'-----------------------------------------------------------------------
'Public Methods
'-----------------------------------------------------------------------

'@Description("Data feed into listbox will be in raw format, so we change then format a column by using this method")
Public Sub ReformatListBoxColumns(ByVal ctrl As MSForms.ListBox, ByVal ColumnIndex As Long, ByVal formatType As DataTypes)
    Dim i As Long
    With ctrl
        For i = 1 To .ListCount - 1
            .List(i, ColumnIndex - 1) = ApplyFormat(.List(i, ColumnIndex - 1), formatType)
        Next i
    End With
End Sub

'@Description("Following function will return string in selected format!")
Public Function ApplyFormat(ByVal InputValue As String, ByVal formatType As DataTypes) As String
    If InputValue <> vbNullString Then InputValue = VBA.Trim$(InputValue)
    
    Select Case formatType
    
        Case DataTypes.TYPE_CURRENCY
            ApplyFormat = AppMethods.FormatStringToCurrency(InputValue)
            
        Case DataTypes.TYPE_DATE
            If InputValue <> vbNullString Then
                ApplyFormat = VBA.Format$(InputValue, GetDateFormat)
            End If
            
    End Select
End Function

'@Description("For approver, we do not want him to edit the record so by using this function we can _
enabled disabled the controls based on user type")
Public Sub FormEditingState(ByVal LockState As Boolean, ParamArray ctrls() As Variant)
    Dim ctrl As Variant
    For Each ctrl In ctrls
        ctrl.Locked = Not LockState
    Next ctrl
    Set ctrl = Nothing
End Sub

'@Description("This method will show popup message")
Public Sub ShowMessage(ByVal message As String, ByVal typeOfMessage As messageType)
    Select Case typeOfMessage
        Case messageType.TYPE_CRITICAL
            MsgBox message, vbCritical, SIGN
        Case messageType.TYPE_INFORMATION
            MsgBox message, vbInformation, SIGN
    End Select
End Sub

'@Description("Following method will check and generate warning labels")
Public Sub UpdateControlAfterValidation(ByVal ctrl As MSForms.Control, ByVal valid As Boolean, Optional warning As WarningType, Optional warningMessage As String = vbNullString)
    With ctrl
        Select Case VBA.TypeName(ctrl)
            Case "TextBox", "ComboBox"
                If valid Then
                    .BorderColor = VALID_FIELD_BORDERCOLOR
                    .ForeColor = VALID_FIELD_FONTCOLOR
                    Call RemoveWarningLabel(ctrl, warning)
                Else
                    .BorderColor = INVALID_FIELD_BORDERCOLOR
                    .ForeColor = INVALID_FIELD_FONTCOLOR
                    If ctrl.Value <> vbNullString Then
                        Call InsertWarningLabel(ctrl, warning, warningMessage)
                    Else
                        Call RemoveWarningLabel(ctrl, warning)
                    End If
                End If
        End Select
    End With
End Sub

'@Description("Following method will set the dimention of the userform")
Public Sub ReDimensionForm()
    With TargetForm
        .Width = formWidth
        .Height = formHeight
    End With
End Sub

'@Description("Hydrate Comboboxes")
Public Sub HydrateComboBox(ByVal cmb As MSForms.ComboBox, ByVal Data As Variant)
    cmb.List = Data
End Sub

'@Description("Change Control's Value and Font Color")
Public Sub ChangeControlProperties(ByVal ctrl As MSForms.Control, ByVal ctrlValue As Variant, Optional ctrlFontColor As Long = vbBlack, Optional ctrlBackColor As Long = vbWhite)
    With ctrl
        Select Case VBA.TypeName(ctrl)
            Case CONTROL_TYPE_TEXTBOX, CONTROL_TYPE_COMBOBOX
                .Value = ctrlValue
            Case CONTROL_TYPE_LABEL
                .Caption = ctrlValue
        End Select
        .ForeColor = ctrlFontColor
        .BackColor = ctrlBackColor
    End With
End Sub

'@Description("Null control's Properties in Bulk")
Public Sub SetStateofControlsToNullState(ParamArray ctrls() As Variant)
    Dim ctrl As Variant
    For Each ctrl In ctrls
        Select Case VBA.TypeName(ctrl)
            Case CONTROL_TYPE_TEXTBOX
                ctrl.Value = vbNullString
            Case CONTROL_TYPE_LABEL
                ctrl.Caption = vbNullString
            Case CONTROL_TYPE_COMBOBOX
                ctrl.Clear
                ctrl.Value = vbNullString
                ctrl.Style = 2
            Case CONTROL_TYPE_LISTBOX
                ctrl.Clear
        End Select
    Next ctrl
End Sub

'@Description("Following Procedure will update set the frames default dimensions")
Public Sub SetDefaultFrameSize(ByVal panelFrame As MSForms.Frame, ByVal nature As String)

    Select Case nature
        Case PANEL_TYPE_SIDE
            RedimensioningOfFramesBasedOnNature panelFrame, PANEL_TYPE_SIDE

        Case PANEL_TYPE_MAIN
            RedimensioningOfFramesBasedOnNature panelFrame, PANEL_TYPE_MAIN

    End Select

End Sub


'@Description("Following Procedure will update make frames visible and also give them predefined dimensions")
Public Sub ActivateFrames(ByVal sidePanelFrame As MSForms.Frame, ByVal mainPanelFrame As MSForms.Frame)
    Dim ctrl As MSForms.Control
    For Each ctrl In TargetForm.Controls
        If VBA.TypeName(ctrl) = CONTROL_TYPE_FRAME Then
            Select Case ctrl.Name
                Case sidePanelFrame.Name
                    If ctrl.Visible = False Then ctrl.Visible = True
                Case mainPanelFrame.Name
                    If ctrl.Visible = False Then ctrl.Visible = True
                Case frameAlwaysOn.Name
                    If IsAlwaysOnFrameVisible = False Then
                        IsAlwaysOnFrameVisible = True
                        RedimensioningOfFramesBasedOnNature frameAlwaysOn, PANEL_TYPE_ALWAYSON
                        ctrl.Visible = True
                    End If
                Case Else
                    If ctrl.Visible = True Then ctrl.Visible = False
            End Select
        End If
    Next ctrl
    Set ctrl = Nothing
End Sub

'-----------------------------------------------------------------------
'Private abstract methods
'-----------------------------------------------------------------------

'@Description("Following method will add label on runtime. if it present then it will only change the visible property of the label")
Private Sub InsertWarningLabel(ByVal ctrl As MSForms.Control, Optional warning As WarningType, Optional warningMessage As String)
    Select Case warning
        Case WarningType.TYPE_NA
            'do nothing
        Case WarningType.TYPE_NUMBERSONLY
            HidingLabel ctrl, CONTROL_NAME_OK
            InsertingLabel ctrl, "TYPE_NUMBERSONLY", "Only numerics are allowed!"
        Case WarningType.TYPE_STRINGSNOTMATCHED
            HidingLabel ctrl, CONTROL_NAME_OK
            InsertingLabel ctrl, "TYPE_STRINGSNOTMATCHED", "Doesn't matched!" & " " & warningMessage
        Case WarningType.TYPE_WRONGPASSWORDPATTERN
            HidingLabel ctrl, CONTROL_NAME_OK
            InsertingLabel ctrl, "TYPE_WRONGPASSWORDPATTERN", "Password must have min. 8 characters, One A-Z, a-z, 0-9 and Special Character"
        Case WarningType.TYPE_FIXEDLENGTHSTRING
            HidingLabel ctrl, CONTROL_NAME_OK
            InsertingLabel ctrl, "TYPE_FIXEDLENGTHSTRING", "String must have minimum of " & warningMessage & " characters"
        Case WarningType.TYPE_CUSTOM
            HidingLabel ctrl, CONTROL_NAME_OK
            InsertingLabel ctrl, "TYPE_CUSTOM", warningMessage
        Case WarningType.TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied
            HidingLabel ctrl, CONTROL_NAME_OK
            InsertingLabel ctrl, "TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied", warningMessage
    End Select
End Sub

'@Description("Following method wiil remove label runtime.")
Private Sub RemoveWarningLabel(ByVal ctrl As MSForms.Control, ByVal ctrlName As String)
    Select Case ctrlName
        Case WarningType.TYPE_NA
            Exit Sub
        Case WarningType.TYPE_NUMBERSONLY
            HidingLabel ctrl, "TYPE_NUMBERSONLY"
        Case WarningType.TYPE_STRINGSNOTMATCHED
            HidingLabel ctrl, "TYPE_STRINGSNOTMATCHED"
        Case WarningType.TYPE_WRONGPASSWORDPATTERN
            HidingLabel ctrl, "TYPE_WRONGPASSWORDPATTERN"
        Case WarningType.TYPE_FIXEDLENGTHSTRING
            HidingLabel ctrl, "TYPE_FIXEDLENGTHSTRING"
        Case WarningType.TYPE_CUSTOM
            HidingLabel ctrl, "TYPE_CUSTOM"
        Case WarningType.TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied
            HidingLabel ctrl, "TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied"
    End Select
    If ctrl.Value <> vbNullString Then
        InsertingLabel ctrl, CONTROL_NAME_OK, SYMBOL_OK, True
    Else
        HidingLabel ctrl, CONTROL_NAME_OK
    End If
End Sub

'@Description("Inserting Warning Label")
Private Sub InsertingLabel(ctrl As MSForms.Control, ByVal ctrlName As String, ByVal message As String, Optional IsItSymbol As Boolean = False)
    Dim TargetTop As Long
    Dim TargetLeft As Long
    Dim TargetFrame As MSForms.Frame
    Dim warningLabel As MSForms.Label
    Dim LabelName As String
    'Setting up Label
    Set TargetFrame = ctrl.Parent
    LabelName = ctrl.Name & "_" & ctrlName
    TargetTop = ctrl.Top
    TargetLeft = ctrl.Left + ctrl.Width + 6
    On Error Resume Next
    Set warningLabel = TargetForm.Controls(LabelName)
    If Err Then
        Set warningLabel = TargetFrame.Controls.Add("Forms.Label.1", LabelName, True)
        Err.Clear
    End If
    On Error GoTo 0
    
    With warningLabel
        .Visible = True
        .Top = TargetTop
        .Left = TargetLeft
        .Caption = IIf(IsItSymbol, message, "* " & message)
        If IsItSymbol Then
            .Height = 18
            .Width = 18
            .ForeColor = COLOR_OF_OKAY
            With .Font
                .Charset = 2
                .Name = "Wingdings"
                .Size = 22
            End With
        Else
            .WordWrap = True
            .AutoSize = True
            .Width = TargetFrame.Width - ctrl.Left - ctrl.Width - 12
            .ForeColor = COLOR_OF_NOT_OKAY
            .Font.Size = 9
        End If
    End With
    Set TargetFrame = Nothing
    Set warningLabel = Nothing
End Sub

'@Description("Removing Warning Labels")
Private Sub HidingLabel(ByVal ctrl As MSForms.Control, ByVal ctrlName As String)
    Dim warningLabel As MSForms.Label
    Dim LabelName As String
    LabelName = ctrl.Name & "_" & ctrlName
    On Error Resume Next
    Set warningLabel = TargetForm.Controls(LabelName)
    warningLabel.Visible = False
    If Err Then Err.Clear
    On Error GoTo 0
    Set warningLabel = Nothing
End Sub

'@Description("Following method will distribute frame dimensions based on their nature")
Private Sub RedimensioningOfFramesBasedOnNature(ByVal ctrl As MSForms.Control, ByVal nature As String)
    If nature = PANEL_TYPE_SIDE Then RedimensionTheFrame ctrl, sideFrameTop, sideFrameLeft, sideFrameWidth, sideFrameHeight
    If nature = PANEL_TYPE_MAIN Then RedimensionTheFrame ctrl, mainFrameTop, mainFrameLeft, mainFrameWidth, mainFrameHeight
    If nature = PANEL_TYPE_ALWAYSON Then RedimensionTheFrame ctrl, alwaysOnTop, alwaysOnLeft, alwaysOnWidth, alwaysOnHeight
End Sub

'@Description("Redim the frame")
Private Sub RedimensionTheFrame(ByVal ctrl As MSForms.Control, ByVal Top As Integer, ByVal Left As Integer, ByVal Width As Integer, ByVal Height As Integer)
    With ctrl
        .Top = Top
        .Left = Left
        .Width = Width
        .Height = Height
    End With
End Sub
