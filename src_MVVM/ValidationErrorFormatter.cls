VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationErrorFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object that describes how a control changes its appearance given a validation error."
'@Folder("MVVM.Validation.ErrorFormatting")
'@ModuleDescription "An object that describes how a control changes its appearance given a validation error."
'@PredeclaredId
'@Exposed
'@IgnoreModule ValueRequiredInspection false positive, stdole.OLE_COLOR
Option Explicit
Implements IValidationErrorFormatter

Private Const DefaultErrorBackColor As Long = &HC0C0FF
Private Const DefaultErrorBorderColor As Long = &HC0
Private Const DefaultErrorBorderWidth As Long = 2
Private Const DefaultErrorForeColor As Long = &HC0

Private Type TState
    Applied As Boolean
    
    FormatBackgroundColor As Boolean
    FormatBorderColor As Boolean
    FormatForeColor As Boolean
    FormatFontBold As Boolean
    
    InitialBackgroundColor As Long
    InitialBorderColor As Long
    InitialBorderWidth As Double
    InitialForeColor As Long
    InitialFontBold As Boolean
    
    ErrorBackgroundColor As Long
    ErrorBorderColor As Long
    ErrorBorderWidth As Double
    ErrorForeColor As Long
    ErrorFontBold As Boolean
    
    TargetOnlyVisibleOnError As Boolean
    PropertyName As String
    
End Type

Private this As TState

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is ValidationErrorFormatter
End Property

'@Description "A builder method to create or configure a formatter that sets the target's background color when applied."
Public Function WithErrorBackgroundColor(Optional ByVal Color As Variant = DefaultErrorBackColor) As ValidationErrorFormatter
Attribute WithErrorBackgroundColor.VB_Description = "A builder method to create or configure a formatter that sets the target's background color when applied."
    Dim result As ValidationErrorFormatter
    If IsDefaultInstance Then
        Set result = New ValidationErrorFormatter
        Set result = result.WithErrorBackgroundColor(Color)
    Else
        Set result = Me
        this.ErrorBackgroundColor = Color
        this.FormatBackgroundColor = True
    End If
    Set WithErrorBackgroundColor = result
End Function

'@Description "A builder method to create or configure a formatter that sets the target's border color when applied."
Public Function WithErrorBorderColor(Optional ByVal Color As Variant = DefaultErrorBorderColor) As ValidationErrorFormatter
Attribute WithErrorBorderColor.VB_Description = "A builder method to create or configure a formatter that sets the target's border color when applied."
    Dim result As ValidationErrorFormatter
    If IsDefaultInstance Then
        Set result = New ValidationErrorFormatter
        Set result = result.WithErrorBorderColor(Color)
    Else
        Set result = Me
        this.ErrorBorderColor = Color
        this.FormatBorderColor = True
    End If
    Set WithErrorBorderColor = result
End Function

'@Description "A builder method to create or configure a formatter that sets the target's border weight when applied (Excel.Range targets only)."
Public Function WithErrorBorderWidth(Optional ByVal Width As Double = DefaultErrorBorderWidth) As ValidationErrorFormatter
Attribute WithErrorBorderWidth.VB_Description = "A builder method to create or configure a formatter that sets the target's border weight when applied (Excel.Range targets only)."
    Dim result As ValidationErrorFormatter
    If IsDefaultInstance Then
        Set result = New ValidationErrorFormatter
        Set result = result.WithErrorBorderWidth(Width)
    Else
        Set result = Me
        this.ErrorBorderWidth = Width
        this.FormatBorderColor = True
    End If
    Set WithErrorBorderWidth = result
End Function

'@Description "A builder method to create or configure a formatter that sets the target's forecolor (i.e. text color) when applied."
Public Function WithErrorForeColor(Optional ByVal Color As Variant = DefaultErrorForeColor) As ValidationErrorFormatter
Attribute WithErrorForeColor.VB_Description = "A builder method to create or configure a formatter that sets the target's forecolor (i.e. text color) when applied."
    Dim result As ValidationErrorFormatter
    If IsDefaultInstance Then
        Set result = New ValidationErrorFormatter
        Set result = result.WithErrorForeColor(Color)
    Else
        Set result = Me
        this.ErrorForeColor = Color
        this.FormatForeColor = True
    End If
    Set WithErrorForeColor = result
End Function

'@Description "A builder method to create or configure a formatter that makes the target's font bold when applied."
Public Function WithErrorBoldFont() As ValidationErrorFormatter
Attribute WithErrorBoldFont.VB_Description = "A builder method to create or configure a formatter that makes the target's font bold when applied."
    Dim result As ValidationErrorFormatter
    If IsDefaultInstance Then
        Set result = New ValidationErrorFormatter
        Set result = result.WithErrorBoldFont
    Else
        Set result = Me
        this.FormatFontBold = True
        this.ErrorFontBold = True
    End If
    Set WithErrorBoldFont = result
End Function

'@Description "A builder method to create or configure a formatter that makes the target only visible when formatter is applied."
Public Function WithTargetOnlyVisibleOnError() As ValidationErrorFormatter
Attribute WithTargetOnlyVisibleOnError.VB_Description = "A builder method to create or configure a formatter that makes the target only visible when formatter is applied."
    Dim result As ValidationErrorFormatter
    If IsDefaultInstance Then
        Set result = New ValidationErrorFormatter
        Set result = result.WithTargetOnlyVisibleOnError
    Else
        Set result = Me
        this.TargetOnlyVisibleOnError = True
    End If
    Set WithTargetOnlyVisibleOnError = result
End Function

Private Sub StoreInitialFormat(ByVal RHS As Object)
    'GuardClauses..GuardDefaultInstance Me, ValidationErrorFormatter, TypeName(Me)
    
    Select Case True
    
        Case TypeOf RHS Is MSForms.TextBox
            Dim TextBoxTarget As MSForms.TextBox
            Set TextBoxTarget = RHS
            this.InitialBackgroundColor = TextBoxTarget.BackColor
            this.InitialBorderColor = TextBoxTarget.BorderColor
            this.InitialFontBold = TextBoxTarget.Font.Bold
            this.InitialForeColor = TextBoxTarget.ForeColor
            
        Case TypeOf RHS Is MSForms.CheckBox 'NOTE: MSForms.OptionButton also matches this interface
            Dim CheckBoxTarget As MSForms.CheckBox
            Set CheckBoxTarget = RHS
            this.InitialFontBold = CheckBoxTarget.Font.Bold
            this.InitialForeColor = CheckBoxTarget.ForeColor
            
        Case TypeOf RHS Is MSForms.Label
            Dim LabelTarget As MSForms.Label
            Set LabelTarget = RHS
            this.InitialBackgroundColor = LabelTarget.BackColor
            this.InitialBorderColor = LabelTarget.BorderColor
            this.InitialFontBold = LabelTarget.Font.Bold
            this.InitialForeColor = LabelTarget.ForeColor
            
        Case TypeOf RHS Is MSForms.Frame
            Dim FrameTarget As MSForms.Frame
            Set FrameTarget = RHS
            this.InitialBackgroundColor = FrameTarget.BackColor
            this.InitialBorderColor = FrameTarget.BorderColor
            this.InitialFontBold = FrameTarget.Font.Bold
            this.InitialForeColor = FrameTarget.ForeColor
        
        Case TypeOf RHS Is MSForms.Image
            Dim ImageTarget As MSForms.Image
            Set ImageTarget = RHS
            this.InitialBackgroundColor = ImageTarget.BackColor
            this.InitialBorderColor = ImageTarget.BorderColor
            
        Case TypeOf RHS Is Excel.Range
            Dim RangeTarget As Excel.Range
            Set RangeTarget = RHS
            this.InitialBackgroundColor = RangeTarget.Interior.Color
            this.InitialBorderColor = RangeTarget.Borders.Color
            this.InitialBorderWidth = RangeTarget.Borders.Weight
            this.InitialFontBold = RangeTarget.Font.Bold
            this.InitialForeColor = RangeTarget.Font.Color
            
        Case Else
            Debug.Print TypeName(Me) & ": Target of type " & TypeName(RHS) & " is not currently supported."
            
    End Select
    
End Sub

Private Sub IValidationErrorFormatter_Apply(ByVal Target As Object, ByVal Message As String)

    If this.Applied Then Exit Sub
    StoreInitialFormat Target
    
    If TypeOf Target Is MSForms.Control Then
        Dim Control As MSForms.Control
        Set Control = Target
        If Not this.TargetOnlyVisibleOnError Then
            Control.ControlTipText = Message
        End If
    End If
    
    Select Case True
    
        Case TypeOf Target Is MSForms.TextBox
            Dim TextBoxTarget As MSForms.TextBox
            Set TextBoxTarget = Target
            If this.FormatBackgroundColor Then TextBoxTarget.BackColor = this.ErrorBackgroundColor
            If this.FormatBorderColor Then TextBoxTarget.BorderColor = this.ErrorBorderColor
            If this.FormatFontBold Then TextBoxTarget.Font.Bold = this.ErrorFontBold
            If this.FormatForeColor Then TextBoxTarget.ForeColor = this.ErrorForeColor
            this.Applied = True
            
        Case TypeOf Target Is MSForms.CheckBox 'NOTE: MSForms.OptionButton also matches this interface
            Dim CheckBoxTarget As MSForms.CheckBox
            Set CheckBoxTarget = Target
            If this.FormatFontBold Then CheckBoxTarget.Font.Bold = this.ErrorFontBold
            If this.FormatForeColor Then CheckBoxTarget.ForeColor = this.ErrorForeColor
            this.Applied = True
            
        Case TypeOf Target Is MSForms.Label
            Dim LabelTarget As MSForms.Label
            Set LabelTarget = Target
            If this.FormatBackgroundColor Then LabelTarget.BackColor = this.ErrorBackgroundColor
            If this.FormatBorderColor Then LabelTarget.BorderColor = this.ErrorBorderColor
            If this.FormatFontBold Then LabelTarget.Font.Bold = this.ErrorFontBold
            If this.FormatForeColor Then LabelTarget.ForeColor = this.ErrorForeColor
            If this.TargetOnlyVisibleOnError Then LabelTarget.Caption = Message
            this.Applied = True
            
        Case TypeOf Target Is MSForms.Frame
            Dim FrameTarget As MSForms.Frame
            Set FrameTarget = Target
            If this.FormatBackgroundColor Then FrameTarget.BackColor = this.ErrorBackgroundColor
            If this.FormatBorderColor Then FrameTarget.BorderColor = this.ErrorBorderColor
            If this.FormatFontBold Then FrameTarget.Font.Bold = this.ErrorFontBold
            If this.FormatForeColor Then FrameTarget.ForeColor = this.ErrorForeColor
            this.Applied = True
            
        Case TypeOf Target Is MSForms.Image
            Dim ImageTarget As MSForms.Image
            Set ImageTarget = Target
            If this.FormatBackgroundColor Then ImageTarget.BackColor = this.ErrorBackgroundColor
            If this.FormatBorderColor Then ImageTarget.BorderColor = this.ErrorBorderColor
            this.Applied = True
            
        Case TypeOf Target Is Excel.Range
            Dim RangeTarget As Excel.Range
            Set RangeTarget = Target
            If this.FormatBackgroundColor Then RangeTarget.Interior.Color = this.ErrorBackgroundColor
            If this.FormatBorderColor Then
                RangeTarget.Borders.Color = this.ErrorBorderColor
                RangeTarget.Borders.Weight = this.ErrorBorderWidth
            End If
            If this.FormatFontBold Then RangeTarget.Font.Bold = this.ErrorFontBold
            If this.FormatForeColor Then RangeTarget.Font.Color = this.ErrorForeColor
    
    End Select
    
    If this.TargetOnlyVisibleOnError And Not Control Is Nothing Then
        Control.Visible = this.Applied
    End If

End Sub


Private Sub IValidationErrorFormatter_Restore(ByVal Target As Object)
    If Not this.Applied Then Exit Sub
    
    If TypeOf Target Is MSForms.Control Then
        Dim Control As MSForms.Control
        Set Control = Target
        Control.ControlTipText = vbNullString
    End If
    
    Select Case True
    
        Case TypeOf Target Is MSForms.TextBox
            Dim TextBoxTarget As MSForms.TextBox
            Set TextBoxTarget = Target
            If this.FormatBackgroundColor Then TextBoxTarget.BackColor = this.InitialBackgroundColor
            If this.FormatBorderColor Then TextBoxTarget.BorderColor = this.InitialBorderColor
            If this.FormatFontBold Then TextBoxTarget.Font.Bold = this.InitialFontBold
            If this.FormatForeColor Then TextBoxTarget.ForeColor = this.InitialForeColor
            
            this.Applied = False
            
        Case TypeOf Target Is MSForms.CheckBox 'NOTE: MSForms.OptionButton also matches this interface
            Dim CheckBoxTarget As MSForms.CheckBox
            Set CheckBoxTarget = Target
            If this.FormatFontBold Then CheckBoxTarget.Font.Bold = this.InitialFontBold
            If this.FormatForeColor Then CheckBoxTarget.ForeColor = this.InitialForeColor
            
            this.Applied = False
            
        Case TypeOf Target Is MSForms.Label
            Dim LabelTarget As MSForms.Label
            Set LabelTarget = Target
            If this.FormatBackgroundColor Then LabelTarget.BackColor = this.InitialBackgroundColor
            If this.FormatBorderColor Then LabelTarget.BorderColor = this.InitialBorderColor
            If this.FormatFontBold Then LabelTarget.Font.Bold = this.InitialFontBold
            If this.FormatForeColor Then LabelTarget.ForeColor = this.InitialForeColor
            
            this.Applied = False
            
        Case TypeOf Target Is MSForms.Frame
            Dim FrameTarget As MSForms.Frame
            Set FrameTarget = Target
            If this.FormatBackgroundColor Then FrameTarget.BackColor = this.InitialBackgroundColor
            If this.FormatBorderColor Then FrameTarget.BorderColor = this.InitialBorderColor
            If this.FormatFontBold Then FrameTarget.Font.Bold = this.InitialFontBold
            If this.FormatForeColor Then FrameTarget.ForeColor = this.InitialForeColor
        
            this.Applied = False
            
        Case TypeOf Target Is MSForms.Image
            Dim ImageTarget As MSForms.Image
            Set ImageTarget = Target
            If this.FormatBackgroundColor Then ImageTarget.BackColor = this.InitialBackgroundColor
            If this.FormatBorderColor Then ImageTarget.BorderColor = this.InitialBorderColor
            
            this.Applied = False
            
        Case TypeOf Target Is Excel.Range
            Dim RangeTarget As Excel.Range
            Set RangeTarget = Target
            If this.FormatBackgroundColor Then RangeTarget.Interior.Color = this.InitialBackgroundColor
            If this.FormatBorderColor Then
                RangeTarget.Borders.Color = this.InitialBorderColor
                RangeTarget.Borders.Weight = this.InitialBorderWidth
            End If
            If this.FormatFontBold Then RangeTarget.Font.Bold = this.InitialFontBold
            If this.FormatForeColor Then RangeTarget.Font.Color = this.InitialForeColor
            
    End Select

    If this.TargetOnlyVisibleOnError And Not Control Is Nothing Then
        Control.Visible = this.Applied
    End If

End Sub
