VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MultiFrameViewExtended"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'By Kamal Bharakhda

'Settings
Const CONTROL_TYPE_FRAME As String = "Frame"
Const CONTROL_TYPE_TEXTBOX As String = "TextBox"
Const CONTROL_TYPE_LABEL As String = "Label"
Const CONTROL_TYPE_COMBOBOX As String = "ComboBox"
Const CONTROL_TYPE_LISTBOX As String = "ListBox"
Const PANEL_TYPE_ALWAYSON As String = "INFO"
Const PANEL_TYPE_SIDE As String = "SIDE"
Const PANEL_TYPE_MAIN As String = "MAIN"

'Public Properties

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

'Public Methods

'Following method will set the dimention of the userform

Public Sub ReDimensionForm()
    With TargetForm
        .Width = formWidth
        .Height = formHeight
    End With
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
            Select Case ctrl.name
                Case sidePanelFrame.name
                    Call RedimensioningOfFramesBasedOnNature(sidePanelFrame, PANEL_TYPE_SIDE)
                    If ctrl.Visible = False Then ctrl.Visible = True
                Case mainPanelFrame.name
                    Call RedimensioningOfFramesBasedOnNature(mainPanelFrame, PANEL_TYPE_MAIN)
                    If ctrl.Visible = False Then ctrl.Visible = True
                Case frameAlwaysOn.name
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

'Private abstract methods

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
