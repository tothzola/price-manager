VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationErrorAdorner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that dynamically decorates a target MSForms.Control object."
'@Folder("MVVM.Validation")
'@ModuleDescription "An object that dynamically decorates a target MSForms.Control object."
'@PredeclaredId
'@Exposed
Option Explicit
Implements IDynamicAdorner
Implements IDisposable

Private Type TState
    Target As MSForms.Control
    TargetFormatter As IValidationErrorFormatter
    
    AdornerLabel As MSForms.Label
    LabelFormatter As IValidationErrorFormatter
    LabelControl As MSForms.Control
    
    AdornerIcon As MSForms.Image
    IconFormatter As IValidationErrorFormatter
    IconControl As MSForms.Control
    
    LabelPosition As IDynamicPosition
    IconPosition As IDynamicPosition
    
    LabelPositionSet As Boolean
    IconPositionSet As Boolean
    Visible As Boolean
    
End Type

Private this As TState

Public Function Create(ByVal Target As MSForms.Control, _
Optional ByVal TargetFormatter As IValidationErrorFormatter = Nothing, _
Optional ByVal AdornerLabel As Variant, _
Optional ByVal LabelFormatter As IValidationErrorFormatter = Nothing, _
Optional ByVal LabelPosition As IDynamicPosition = Nothing, _
Optional ByVal AdornerIcon As Variant, _
Optional ByVal IconFormatter As IValidationErrorFormatter = Nothing, _
Optional ByVal IconPosition As IDynamicPosition = Nothing) As IDynamicAdorner
    
    'GuardClauses.GuardNonDefaultInstance Me, ValidationErrorAdorner
    'GuardClauses.GuardNullReference Target, TypeName(Me)
    
    Dim result As ValidationErrorAdorner
    Set result = New ValidationErrorAdorner
    
    InitTarget result, Target, TargetFormatter
    If IsMissing(AdornerLabel) Then
        Set AdornerLabel = Nothing
        InitLabel result, AdornerLabel, LabelFormatter, LabelPosition
    ElseIf IsObject(AdornerLabel) Then
        If Not AdornerLabel Is Nothing Then
            InitLabel result, AdornerLabel, LabelFormatter, LabelPosition
        End If
    End If
    
    If IsMissing(AdornerIcon) Then
        Set AdornerIcon = Nothing
        InitIcon result, AdornerIcon, IconFormatter, IconPosition
    ElseIf IsObject(AdornerIcon) Then
        If Not AdornerIcon Is Nothing Then
            Set AdornerIcon = AdornerIcon
            InitIcon result, AdornerIcon, IconFormatter, IconPosition
        End If
    End If
    
    Set Create = result
    
End Function

Private Sub InitTarget(ByVal result As ValidationErrorAdorner, ByVal Target As Object, ByVal Formatter As IValidationErrorFormatter)
    Set result.Target = Target
    If Formatter Is Nothing Then
        Set result.TargetFormatter = DefaultTargetFormatter
    Else
        Set result.TargetFormatter = Formatter
    End If
End Sub

Private Sub InitLabel(ByVal result As ValidationErrorAdorner, ByRef AdornerLabel As Variant, ByRef Formatter As IValidationErrorFormatter, ByRef Position As IDynamicPosition)
    If Position Is Nothing Then Set Position = DefaultLabelPosition
    
    If AdornerLabel Is Nothing Then
        Dim Parent As MSForms.Controls
        Set Parent = result.Target.Parent.Controls
        
        Set AdornerLabel = Parent.Add(FormsProgID.LabelProgId, Visible:=False)
        
        AdornerLabel.BackStyle = fmBackStyleTransparent
        AdornerLabel.WordWrap = False
        AdornerLabel.AutoSize = True
        AdornerLabel.Font.Size = result.Target.Parent.Font.Size - 2
                        
    End If
    
    PositionAdornerControlX result.Target, AdornerLabel, Position
    PositionAdornerControlY result.Target, AdornerLabel, Position
    
    Set result.AdornerLabel = AdornerLabel
    Set result.AdornerLabelPosition = Position
    If Formatter Is Nothing Then
        Set result.AdornerLabelFormatter = DefaultLabelFormatter
    Else
        Set result.AdornerLabelFormatter = Formatter
    End If
End Sub

Private Sub InitIcon(ByVal result As ValidationErrorAdorner, ByRef AdornerIcon As Variant, ByRef Formatter As IValidationErrorFormatter, ByRef Position As IDynamicPosition)
    If Position Is Nothing Then Set Position = DefaultIconPosition
    
    If AdornerIcon Is Nothing Then
        Dim Parent As MSForms.Controls
        Set Parent = result.Target.Parent.Controls
        
        Set AdornerIcon = Parent.Add(FormsProgID.ImageProgId, Visible:=False)
        Set AdornerIcon.Picture = Resources.ValidationErrorIcon.Picture
        
        AdornerIcon.BackStyle = fmBackStyleOpaque
        AdornerIcon.BorderStyle = fmBorderStyleSingle
        AdornerIcon.PictureSizeMode = fmPictureSizeModeClip
        AdornerIcon.PictureAlignment = fmPictureAlignmentCenter
        AdornerIcon.PictureTiling = False
        AdornerIcon.AutoSize = True
        
    End If
    
    PositionAdornerControlX result.Target, AdornerIcon, Position
    PositionAdornerControlY result.Target, AdornerIcon, Position
    
    Set result.AdornerIcon = AdornerIcon
    Set result.AdornerIconPosition = Position
    If Formatter Is Nothing And Not AdornerIcon Is Nothing Then
        Set result.AdornerIconFormatter = DefaultIconFormatter
    Else
        Set result.AdornerIconFormatter = Formatter
    End If
End Sub

Private Property Get DefaultTargetFormatter() As IValidationErrorFormatter
    Set DefaultTargetFormatter = ValidationErrorFormatter _
        .WithErrorBackgroundColor _
        .WithErrorBoldFont _
        .WithErrorBorderColor _
        .WithErrorForeColor
End Property

Private Property Get DefaultLabelFormatter() As IValidationErrorFormatter
    Set DefaultLabelFormatter = ValidationErrorFormatter _
        .WithErrorBoldFont _
        .WithErrorForeColor _
        .WithTargetOnlyVisibleOnError
End Property

Private Property Get DefaultLabelPosition() As IDynamicPosition
    Set DefaultLabelPosition = DynamicControlPosition.Create(RelativePosition.BelowRight, 3, 0)
End Property

Private Property Get DefaultIconFormatter() As IValidationErrorFormatter
    Set DefaultIconFormatter = ValidationErrorFormatter _
        .WithErrorBorderColor _
        .WithTargetOnlyVisibleOnError
End Property

Private Property Get DefaultIconPosition() As IDynamicPosition
    Set DefaultIconPosition = DynamicControlPosition.Create(RelativePosition.InsideRight, 3)
End Property

Friend Property Get Target() As Object
    Set Target = this.Target
End Property

Friend Property Set Target(ByVal RHS As Object)
    'GuardClauses.GuardDefaultInstance Me, ValidationErrorAdorner
    'GuardClauses.GuardDoubleInitialization This.Target, TypeName(Me)
    Set this.Target = RHS
End Property

Friend Property Get TargetFormatter() As IValidationErrorFormatter
    Set TargetFormatter = this.TargetFormatter
End Property

Friend Property Set TargetFormatter(ByVal RHS As IValidationErrorFormatter)
    Set this.TargetFormatter = RHS
End Property

Public Property Get Visible() As Boolean
    Visible = this.Visible
End Property

Private Sub PositionAdornerControlX(ByVal Target As MSForms.Control, ByVal Adorner As MSForms.Control, ByVal Position As IDynamicPosition)
    Select Case Position.Position
    
        Case RelativePosition.AboveRight, _
             RelativePosition.BelowRight, _
             RelativePosition.InsideRight, _
             RelativePosition.Default
            
            TrySetTextAlign Adorner, fmTextAlignRight
            Adorner.Left = Target.Left + Target.Width - Adorner.Width - Position.Margin(RightSide)
            
        Case RelativePosition.AboveLeft, _
             RelativePosition.BelowLeft, _
             RelativePosition.InsideLeft
            
            TrySetTextAlign Adorner, fmTextAlignLeft
            Adorner.Left = Target.Left + Position.Margin(LeftSide)
            
    End Select
End Sub

Private Sub TrySetTextAlign(ByVal Target As Object, ByVal Value As MSForms.fmTextAlign)
    'late-bound assignment will fail if Target doesn't have a TextAlign property.
    On Error Resume Next
    'if execution stops here, go Tools ~> Options ~> General ~> Error Trapping,
    'and make sure "Break on unhandled errors" is selected.
    'only use "Break on all errors" for debugging & diagnostics.
    Target.TextAlign = Value
    On Error GoTo 0
End Sub

Private Sub PositionAdornerControlY(ByVal Target As MSForms.Control, ByVal Adorner As MSForms.Control, ByVal Position As IDynamicPosition)
    Select Case Position.Position
    
        Case RelativePosition.AboveLeft, _
             RelativePosition.AboveRight
            
            Adorner.Top = Target.Top - Adorner.Height - Position.Margin(BottomSide)
            
        Case RelativePosition.InsideLeft, _
             RelativePosition.InsideRight
            
            Adorner.Top = Target.Top + Position.Margin(TopSide)
            
        Case RelativePosition.BelowLeft, _
             RelativePosition.BelowRight, _
             RelativePosition.Default
            
            Adorner.Top = Target.Top + Target.Height + Position.Margin(TopSide)
            
    End Select
End Sub

Friend Property Get AdornerLabel() As MSForms.Label
    Set AdornerLabel = this.AdornerLabel
End Property

Friend Property Set AdornerLabel(ByVal RHS As MSForms.Label)
    Set this.AdornerLabel = RHS
    Set this.LabelControl = RHS
End Property

Friend Property Get AdornerLabelFormatter() As IValidationErrorFormatter
    Set AdornerLabelFormatter = this.LabelFormatter
End Property

Friend Property Set AdornerLabelFormatter(ByVal RHS As IValidationErrorFormatter)
    Set this.LabelFormatter = RHS
End Property

Friend Property Get AdornerLabelPosition() As IDynamicPosition
    Set AdornerLabelPosition = this.LabelPosition
End Property

Friend Property Set AdornerLabelPosition(ByVal RHS As IDynamicPosition)
    Set this.LabelPosition = RHS
End Property

Friend Property Get AdornerIcon() As MSForms.Image
    Set AdornerIcon = this.AdornerIcon
End Property

Friend Property Set AdornerIcon(ByVal RHS As MSForms.Image)
    Set this.AdornerIcon = RHS
    Set this.IconControl = RHS
End Property

Friend Property Get AdornerIconFormatter() As IValidationErrorFormatter
    Set AdornerIconFormatter = this.IconFormatter
End Property

Friend Property Set AdornerIconFormatter(ByVal RHS As IValidationErrorFormatter)
    Set this.IconFormatter = RHS
End Property

Friend Property Get AdornerIconPosition() As IDynamicPosition
    Set AdornerIconPosition = this.IconPosition
End Property

Friend Property Set AdornerIconPosition(ByVal RHS As IDynamicPosition)
    Set this.IconPosition = RHS
End Property

Private Sub IDisposable_Dispose()
    Set this.Target = Nothing
    Set this.AdornerIcon = Nothing
    Set this.AdornerLabel = Nothing
    Set this.IconControl = Nothing
    Set this.LabelControl = Nothing
End Sub

Private Sub IDynamicAdorner_Hide()
    If Not this.Visible Then Exit Sub
    'GuardClauses.GuardDefaultInstance Me, ValidationErrorAdorner
    'GuardClauses.GuardNullReference This.TargetFormatter, TypeName(Me), "TargetFormatter is not set for this instance."
    
    this.TargetFormatter.Restore this.Target
    If Not this.IconControl Is Nothing Then this.IconFormatter.Restore this.IconControl
    If Not this.LabelControl Is Nothing Then this.LabelFormatter.Restore this.LabelControl

    this.Visible = False
End Sub

Private Sub IDynamicAdorner_Show(ByVal Message As String)
    'If This.Visible Then Exit Sub '<~ what if we just wanted to update the message?
    'GuardClauses.GuardDefaultInstance Me, ValidationErrorAdorner
    'GuardClauses.GuardNullReference This.TargetFormatter, TypeName(Me), "TargetFormatter is not set for this instance."
    
    this.TargetFormatter.Apply this.Target, Message
    If Not this.LabelControl Is Nothing Then ShowAdornerLabel Message
    If Not this.IconControl Is Nothing Then ShowAdornerIcon Message
    
    this.Visible = True
End Sub

Private Sub ShowAdornerLabel(ByVal Message As String)
    this.LabelFormatter.Apply this.LabelControl, Message
    If Not this.LabelPositionSet Then
        PositionAdornerControlX this.Target, this.AdornerLabel, this.LabelPosition
        PositionAdornerControlY this.Target, this.AdornerLabel, this.LabelPosition
        this.LabelPositionSet = True
    End If
End Sub

Private Sub ShowAdornerIcon(ByVal Message As String)
    this.IconFormatter.Apply this.IconControl, Message
    If Not this.IconPositionSet Then
        PositionAdornerControlX this.Target, this.AdornerIcon, this.IconPosition
        PositionAdornerControlY this.Target, this.AdornerIcon, this.IconPosition
        this.IconPositionSet = True
    End If
End Sub


