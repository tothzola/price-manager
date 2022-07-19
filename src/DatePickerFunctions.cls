VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DatePickerFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule EncapsulatePublicField
'@Folder("PriceApproval.View")
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
Public WithEvents LabelBackground As MSForms.Label
Attribute LabelBackground.VB_VarHelpID = -1

Private parentFrame As MSForms.Frame

Private Sub LabelBackground_Click()

    Set parentFrame = LabelBackground.Parent
    
    If Not parentFrame Is Nothing Then
        Dim monthsSelectorValue As Long
        monthsSelectorValue = parentFrame.MonthsSelector.Value
    
        Dim labelForeColor As Long
        labelForeColor = (LabelBackground.ForeColor = &H80000010)
    
        Dim labelTitelCaptionUpper As Long
        labelTitelCaptionUpper = (LabelBackground.Caption > 20)
    
        Dim labelTitelCaptionLower As Long
        labelTitelCaptionLower = (LabelBackground.Caption < 18)
    
        Dim calibrate As Long
        calibrate = monthsSelectorValue - labelForeColor * labelTitelCaptionUpper + labelForeColor * labelTitelCaptionLower
        
        Dim selectedDate As Date
        selectedDate = VBA.DateSerial(VBA.Year(VBA.Date), VBA.Month(VBA.Date) + calibrate, LabelBackground.Caption)
        
        parentFrame.Parent.Controls(parentFrame.Tag).Text = selectedDate
        parentFrame.Visible = False
        
    End If
    
End Sub

Private Sub LabelBackground_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    
    Dim j As Integer
    For j = 1 To 42
        Set parentFrame = LabelBackground.Parent
        '@Ignore UnassignedVariableUsage
        If Not parentFrame Is Nothing Then parentFrame.Controls("dpLabel" & j).BackStyle = VBA.Abs(j = VBA.Val(VBA.Mid$(LabelBackground.Name, 8)))
    Next
    
End Sub

Private Sub Class_Terminate()
    Set parentFrame = Nothing
End Sub
