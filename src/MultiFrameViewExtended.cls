VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MultiFrameViewExtended"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "View"
Option Explicit
'By Kamal Bharakhda

'-----------------------------------------------------------------------
'Settings
'-----------------------------------------------------------------------

Const CONTROL_TYPE_FRAME As String = "Frame"
Const CONTROL_TYPE_TEXTBOX As String = "TextBox"
Const CONTROL_TYPE_LABEL As String = "Label"
Const CONTROL_TYPE_COMBOBOX As String = "ComboBox"
Const CONTROL_TYPE_LISTBOX As String = "ListBox"
Const PANEL_TYPE_ALWAYSON As String = "INFO"
Const PANEL_TYPE_SIDE As String = "SIDE"
Const PANEL_TYPE_MAIN As String = "MAIN"
Const SYMBOL_OK As String = "ü"
Const CONTROL_NAME_OK As String = "OK"
Const Error_NonNumericalInput As Long = 1

'Colors
Const INVALID_FIELD_BACKCOLOR As Long = COLOR_OF_NOT_OKAY
Const INVALID_FIELD_FONTCOLOR As Long = vbWhite
Const VALID_FIELD_BACKCOLOR As Long = vbWhite
Const VALID_FIELD_FONTCOLOR As Long = vbBlack

'-----------------------------------------------------------------------
'Public Properties
'-----------------------------------------------------------------------

'Userform Properties
Public TargetForm As Object 'Instance of User Form/VIEW
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

'Data feed into listbox will be in raw format, so we change then format a column by using this method

Public Sub ReformatListBoxColumns(ByVal ctrl As MSForms.ListBox, ByVal columnIndex As Long, ByVal formatType As DataTypes)
    Dim I As Long
    With ctrl
        For I = 1 To .ListCount - 1 Step 1
            .List(I, columnIndex - 1) = ApplyFormat(.List(I, columnIndex - 1), formatType)
        Next I
    End With
End Sub

'Following function will return string in selected format!

Public Function ApplyFormat(ByVal inputValue As String, ByVal formatType As DataTypes) As String
    If inputValue <> vbNullString Then inputValue = VBA.Trim(inputValue)
    Select Case formatType
        Case DataTypes.TYPE_CURRENCY
            ApplyFormat = FormatStringToCurrency(inputValue)
        Case DataTypes.TYPE_DATE
            ApplyFormat = Format(inputValue, DATEFORMAT_FRONTEND)
    End Select
End Function

'For approver, we do not want him to edit the record so by using this function we can enabled disabled the controls based on user type

Public Sub FormEditingState(ByVal LockState As Boolean, ParamArray ctrls() As Variant)
    Dim ctrl As Variant
    For Each ctrl In ctrls
        ctrl.Enabled = LockState
    Next ctrl
    Set ctrl = Nothing
End Sub

'this method will show popup message

Public Sub ShowMessage(ByVal message As String, ByVal typeOfMessage As messageType)
    Select Case typeOfMessage
        Case messageType.TYPE_CRITICAL
            MsgBox message, vbCritical, SIGN
        Case messageType.TYPE_INFORMATION
            MsgBox message, vbInformation, SIGN
    End Select
End Sub

'Following method will check and generate warning labels

Public Sub UpdateControlAfterValidation(ByVal ctrl As MSForms.Control, ByVal valid As Boolean, Optional warning As WarningType, Optional warningMessage As String = vbNullString)
    With ctrl
        Select Case VBA.TypeName(ctrl)
            Case "TextBox", "ComboBox"
                If valid Then
                    .BackColor = VALID_FIELD_BACKCOLOR
                    .ForeColor = VALID_FIELD_FONTCOLOR
                    Call RemoveWarningLabel(ctrl, warning)
                Else
                    .BackColor = INVALID_FIELD_BACKCOLOR
                    .ForeColor = INVALID_FIELD_FONTCOLOR
                    Call InsertWarningLabel(ctrl, warning, warningMessage)
                End If
        End Select
    End With
End Sub

'Following method will set the dimention of the userform

Public Sub ReDimensionForm()
    With TargetForm
        .Width = formWidth
        .Height = formHeight
    End With
End Sub

'Hydrate Comboboxes

Public Sub HydrateComboBox(ByVal cmb As MSForms.ComboBox, ByVal data As Variant)
    cmb.List = data
End Sub

'Change Control's Value and Font Color

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

'Null control's Properties in Bulk

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

'Following Procedure will update make frames visible and also give them predefined dimensions

Public Sub ActivateFrames(ByVal sidePanelFrame As MSForms.Frame, ByVal mainPanelFrame As MSForms.Frame)
    Dim ctrl As MSForms.Control
    For Each ctrl In TargetForm.Controls
        If VBA.TypeName(ctrl) = CONTROL_TYPE_FRAME Then
            Select Case ctrl.Name
                Case sidePanelFrame.Name
                    Call RedimensioningOfFramesBasedOnNature(sidePanelFrame, PANEL_TYPE_SIDE)
                    If ctrl.Visible = False Then ctrl.Visible = True
                Case mainPanelFrame.Name
                    Call RedimensioningOfFramesBasedOnNature(mainPanelFrame, PANEL_TYPE_MAIN)
                    If ctrl.Visible = False Then ctrl.Visible = True
                Case frameAlwaysOn.Name
                    If IsAlwaysOnFrameVisible = False Then
                        IsAlwaysOnFrameVisible = True
                        Call RedimensioningOfFramesBasedOnNature(frameAlwaysOn, PANEL_TYPE_ALWAYSON)
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

'Following method will add label on runtime. if it present then it will only change the visible property of the label

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

'Following method wiil remove label runtime.

Private Sub RemoveWarningLabel(ByVal ctrl As MSForms.Control, ByVal ctrlName As String)
    Select Case ctrlName
        Case WarningType.TYPE_NUMBERSONLY
            HidingLabel ctrl, "TYPE_NUMBERSONLY"
            InsertingLabel ctrl, CONTROL_NAME_OK, SYMBOL_OK, True
        Case WarningType.TYPE_STRINGSNOTMATCHED
            HidingLabel ctrl, "TYPE_STRINGSNOTMATCHED"
            InsertingLabel ctrl, CONTROL_NAME_OK, SYMBOL_OK, True
        Case WarningType.TYPE_WRONGPASSWORDPATTERN
            HidingLabel ctrl, "TYPE_WRONGPASSWORDPATTERN"
            InsertingLabel ctrl, CONTROL_NAME_OK, SYMBOL_OK, True
        Case WarningType.TYPE_FIXEDLENGTHSTRING
            HidingLabel ctrl, "TYPE_FIXEDLENGTHSTRING"
            InsertingLabel ctrl, CONTROL_NAME_OK, SYMBOL_OK, True
        Case WarningType.TYPE_CUSTOM
            HidingLabel ctrl, "TYPE_CUSTOM"
            InsertingLabel ctrl, CONTROL_NAME_OK, SYMBOL_OK, True
        Case WarningType.TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied
            HidingLabel ctrl, "TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied"
            If ctrl.Value <> vbNullString Then
                InsertingLabel ctrl, CONTROL_NAME_OK, SYMBOL_OK, True
            Else
                HidingLabel ctrl, CONTROL_NAME_OK
            End If
    End Select
End Sub

'Inserting Warning Label

Private Sub InsertingLabel(ctrl As MSForms.Control, ByVal ctrlName As String, ByVal message As String, Optional IsItSymbol As Boolean = False)
    Dim TargetTop As Long
    Dim TargetLeft As Long
    Dim TargetFrame As MSForms.Frame
    Dim warningLabel As MSForms.Label
    Dim LabelName As String
    'Setting up Label
    Set TargetFrame = ctrl.parent
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

'Removing Warning Labels

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

'Following method will distribute frame dimensions based on their nature

Private Sub RedimensioningOfFramesBasedOnNature(ByVal ctrl As MSForms.Control, ByVal nature As String)
    If nature = PANEL_TYPE_SIDE Then RedimensionTheFrame ctrl, sideFrameTop, sideFrameLeft, sideFrameWidth, sideFrameHeight
    If nature = PANEL_TYPE_MAIN Then RedimensionTheFrame ctrl, mainFrameTop, mainFrameLeft, mainFrameWidth, mainFrameHeight
    If nature = PANEL_TYPE_ALWAYSON Then RedimensionTheFrame ctrl, alwaysOnTop, alwaysOnLeft, alwaysOnWidth, alwaysOnHeight
End Sub

'redim the frame

Private Sub RedimensionTheFrame(ByVal ctrl As MSForms.Control, ByVal Top As Integer, ByVal Left As Integer, ByVal Width As Integer, ByVal Height As Integer)
    With ctrl
        .Top = Top
        .Left = Left
        .Width = Width
        .Height = Height
    End With
End Sub

'Change text format to #,###.00 Currency Format

Public Function FormatStringToCurrency(ByVal inputValue As String, Optional Prefix As String = vbNullString) As String
    
    'Variable Declaration
    Dim IsDecimalSeperatorUsed  As Boolean
    Dim objRegEx                As Object
    Dim ThousandSeperator       As String
    Dim DecimalSeperator        As String
    Dim DecimalPart             As String
    Dim WholeNumberPart         As String
    Dim ArrDecimalNumber()      As String
    Dim tmp                     As String
    Dim LenInputValue           As Long
    Dim LenDecimalPart          As Long
    Dim I                       As Long
    Dim J                       As Long
        
    'error handeling
    On Error GoTo EhhErr
    
    'Validation Checks
    
        'Check 1 : On null value
        If inputValue = vbNullString Then GoTo CleanExit
        
        'Check 2 : If Non-Numerical Value entered!
        If Prefix <> vbNullString Then
            inputValue = VBA.Replace(inputValue, Prefix & " ", vbNullString)
        End If
        If Not VBA.IsNumeric(inputValue) Then
            Err.Raise _
                Number:=Error_NonNumericalInput + 513, _
                Description:="Non-Numerical character(s) are not allowed!"
        End If
        
    'Get Sepearator String
    ThousandSeperator = Application.ThousandsSeparator
    DecimalSeperator = Application.DecimalSeparator
    
    'check if decimal seperator has been used or not ?
    If VBA.InStr(1, inputValue, DecimalSeperator) > 0 Then
        'Get number parts with Position shifter
        ArrDecimalNumber = VBA.Split(inputValue, DecimalSeperator)
        LenDecimalPart = VBA.Len(ArrDecimalNumber(1))
        If LenDecimalPart = 2 Then
            WholeNumberPart = ArrDecimalNumber(0)
            DecimalPart = ArrDecimalNumber(1)
        ElseIf LenDecimalPart = 3 Then
            If ArrDecimalNumber(0) = 0 Then
                WholeNumberPart = VBA.Left(ArrDecimalNumber(1), 1)
            Else
                WholeNumberPart = ArrDecimalNumber(0) & VBA.Left(ArrDecimalNumber(1), 1)
            End If
            DecimalPart = VBA.Right(ArrDecimalNumber(1), 2)
        ElseIf LenDecimalPart = 1 Then
            FormatStringToCurrency = vbNullString
            GoTo CleanExit
        End If
    Else
        'Length of Input String
        LenInputValue = VBA.Len(inputValue)
        'Get whole number parts
        If LenInputValue <= 2 Then
            WholeNumberPart = 0
        Else
            WholeNumberPart = VBA.Left(inputValue, LenInputValue - 2)
        End If
        'get decimal number part
        If LenInputValue > 1 Then
            DecimalPart = VBA.Right(inputValue, 2)
        Else
            DecimalPart = "0" & inputValue
        End If
    End If
    
    'clean number parts
    Set objRegEx = CreateObject("vbscript.regexp")
    With objRegEx
        .Global = True
        .Pattern = "[^\d]+"
        WholeNumberPart = .Replace(WholeNumberPart, vbNullString)
        DecimalPart = .Replace(DecimalPart, vbNullString)
    End With
    
    'insert Thousand Seperator to whole number part!
    tmp = vbNullString
    If VBA.Len(WholeNumberPart) > 3 Then
        J = 1
        For I = VBA.Len(WholeNumberPart) To 1 Step -1
            If J = 3 Then
                tmp = ThousandSeperator & VBA.Mid(WholeNumberPart, I, 1) & tmp
                J = 0
            Else
                tmp = VBA.Mid(WholeNumberPart, I, 1) & tmp
            End If
            J = J + 1
        Next I
        'Clean whole number part
        If VBA.Left(tmp, 1) = ThousandSeperator Then
            tmp = VBA.Right(tmp, VBA.Len(tmp) - 1)
            tmp = VBA.Trim(tmp)
        End If
        'Submission!
        WholeNumberPart = tmp
    End If
    
    'Joining String
    If Prefix = vbNullString Then
        FormatStringToCurrency = WholeNumberPart & DecimalSeperator & DecimalPart
    Else
        FormatStringToCurrency = Prefix & " " & WholeNumberPart & DecimalSeperator & DecimalPart
    End If
    
CleanExit:
    Set objRegEx = Nothing
    Exit Function
    
EhhErr:
    MsgBox "Error: " & Err.Number & vbNewLine & vbNewLine & Err.Description, vbCritical, "Kamal Bharakhda"
    FormatStringToCurrency = vbNullString
    Err.Clear
    Resume CleanExit
    
End Function
