VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule FunctionReturnValueAlwaysDiscarded
'@Folder("Ribbon.Impl")
'@PredeclaredId
'@Exposed
Option Explicit

Implements IRibbon

#Const LateBind = TestMode

Private Const APP_VERSION As Double = 1.4

Private Type TManager
    Ribbon As Office.IRibbonUI
    Invalidated As Boolean

End Type

Private this As TManager

Public Property Get Ribbon() As Office.IRibbonUI
    Set Ribbon = this.Ribbon
End Property

Public Property Set Ribbon(ByVal RHS As Office.IRibbonUI)
    Set this.Ribbon = RHS
End Property

Public Property Get Invalidated() As Boolean
    Invalidated = this.Invalidated
End Property

Public Function Create(ByVal RibbonUI As Office.IRibbonUI) As IRibbon

    Guard.NullReference RibbonUI
    
    With New RibbonManager
        Set .Ribbon = RibbonUI
        
        Set Create = .Self
    End With
    
End Function

Public Function Self() As IRibbon
    Set Self = Me
End Function

Private Sub Class_Terminate()
    Set this.Ribbon = Nothing
    #If LateBind Then
        Debug.Print TypeName(Me) & " is terminating"
    #End If
End Sub

Private Sub IRibbon_Invalidate()
    Ribbon.Invalidate
    this.Invalidated = True
End Sub

'@Description "Determine the text to go along with your Tab, Groups, and Buttons"
Private Function IRibbon_GetLabel(ByVal Identifier As String, ByRef outLabel As Variant) As Variant
Attribute IRibbon_GetLabel.VB_Description = "Determine the text to go along with your Tab, Groups, and Buttons"

    Select Case Identifier
    
    Case Is = "PriceApprovalTab": outLabel = "Price Approval Manager"
    Case Is = "PriceApprovalGroupA": outLabel = "App"
    Case Is = "ButtonA_01": outLabel = "Price Manager"
    Case Is = "ButtonA_02": outLabel = "Button Label"
    Case Is = "ButtonA_03": outLabel = "Button Label"
    Case Is = "ButtonA_04": outLabel = "Button Label"
    Case Is = "ButtonA_05": outLabel = "Button Label"
    
    Case Is = "PriceApprovalGroupB": outLabel = "Contact"
    Case Is = "ButtonB_01": outLabel = "Send Feedback"
    Case Is = "ButtonB_02": outLabel = "Button Label"
    Case Is = "ButtonB_03": outLabel = "Button Label"
    Case Is = "ButtonB_04": outLabel = "Button Label"
    Case Is = "ButtonB_05": outLabel = "Button Label"

    Case Is = "PriceApprovalGroupC": outLabel = "Help"
    Case Is = "LabelC_01": outLabel = "Version: " & APP_VERSION
    Case Is = "ButtonC_01": outLabel = "Check for Updates..."
    
    Case Else: outLabel = "Button Label"
    
    End Select

    IRibbon_GetLabel = outLabel
    
End Function

'@Description "Tell each button which macro subroutine to run when clicked"
Private Sub IRibbon_OnAction(ByVal Control As Office.IRibbonControl, Optional ByRef index As Integer)
Attribute IRibbon_OnAction.VB_Description = "Tell each button which macro subroutine to run when clicked"
    
    If Control.ID Like "Button*" Then
    
        Select Case Control.ID
        
        'App
        Case Is = "ButtonA_01": modMain.MainPAM
        
        'Feedback
        Case Is = "ButtonB_01": SendEmail.EmailFeedback

        'Help
        Case Is = "ButtonC_01": MsgBox "Not Implementd"
        
        End Select
        
    End If
    
End Sub

'@Description "Determine the Image to go along with your Buttons"
Private Function IRibbon_GetImage(ByVal Identifier As String, ByRef outImage As Variant) As Variant
Attribute IRibbon_GetImage.VB_Description = "Determine the Image to go along with your Buttons"
    
    Select Case Identifier
    
    Case Is = "ButtonA_01": outImage = "HelpGettingStarted"
    Case Is = "ButtonA_02": outImage = "ObjectPictureFill"
    Case Is = "ButtonA_03": outImage = "ObjectPictureFill"
    Case Is = "ButtonA_04": outImage = "ObjectPictureFill"
    Case Is = "ButtonA_05": outImage = "ObjectPictureFill"
    
    Case Is = "ButtonB_01": outImage = "PublishToWeb"
    Case Is = "ButtonB_02": outImage = "ObjectPictureFill"
    Case Is = "ButtonB_03": outImage = "ObjectPictureFill"
    Case Is = "ButtonB_04": outImage = "ObjectPictureFill"
    Case Is = "ButtonB_05": outImage = "ObjectPictureFill"
    
    Case Is = "ButtonC_01": outImage = "RefreshWebView"
    
    Case Else: outImage = "ObjectPictureFill"
    
    End Select
    
    IRibbon_GetImage = outImage
    
End Function

'@Description "Determine if the button size is large or small"
Private Function IRibbon_GetSize(ByVal Identifier As String, ByRef outSize As Variant) As Variant
Attribute IRibbon_GetSize.VB_Description = "Determine if the button size is large or small"

    Const NORMAL As Integer = 0
    Const LARGE As Integer = 1
    
    Select Case Identifier
    
    Case Is = "ButtonA_01": outSize = LARGE
    Case Is = "ButtonA_02": outSize = NORMAL
    Case Is = "ButtonA_03": outSize = NORMAL
    Case Is = "ButtonA_04": outSize = NORMAL
    Case Is = "ButtonA_05": outSize = NORMAL
    
    Case Is = "ButtonB_01": outSize = LARGE
    Case Is = "ButtonB_02": outSize = NORMAL
    Case Is = "ButtonB_03": outSize = NORMAL
    Case Is = "ButtonB_04": outSize = NORMAL
    Case Is = "ButtonB_05": outSize = NORMAL
    
    Case Is = "ButtonC_01": outSize = LARGE
    
    Case Else: outSize = NORMAL
    
    End Select
    
    IRibbon_GetSize = outSize
    
End Function

'@Description "Display a specific macro description when the mouse hovers over a button"
Private Function IRibbon_GetScreenTip(ByVal Identifier As String, ByRef outTipp As Variant) As Variant
Attribute IRibbon_GetScreenTip.VB_Description = "Display a specific macro description when the mouse hovers over a button"
    
    Select Case Identifier
    
    Case Is = "ButtonA_01": outTipp = "Click to Open Application"
    Case Is = "ButtonA_02": outTipp = "Description"
    Case Is = "ButtonA_03": outTipp = "Description"
    Case Is = "ButtonA_04": outTipp = "Description"
    Case Is = "ButtonA_05": outTipp = "Description"
    
    Case Is = "ButtonB_01": outTipp = "Send Feedback to Developers"
    Case Is = "ButtonB_02": outTipp = "Description"
    Case Is = "ButtonB_03": outTipp = "Description"
    Case Is = "ButtonB_04": outTipp = "Description"
    Case Is = "ButtonB_05": outTipp = "Description"
    
    Case Is = "LabelC_01": outTipp = "Actual version installed on this machine"
    Case Is = "ButtonC_01": outTipp = "Click"
    
    Case Else: outTipp = "Description"
    
    End Select
    
    IRibbon_GetScreenTip = outTipp
    
End Function

'@Description "Show/Hide buttons based on how many you need"
Private Function IRibbon_IsVisible(ByVal Control As Office.IRibbonControl, ByRef outBoolean As Variant) As Variant
Attribute IRibbon_IsVisible.VB_Description = "Show/Hide buttons based on how many you need"

    Select Case Control.ID
    
    Case Is = "PriceApprovalTab": outBoolean = True
    Case Is = "PriceApprovalGroupA": outBoolean = True
    Case Is = "ButtonA_01": outBoolean = True
    Case Is = "ButtonA_02": outBoolean = False
    Case Is = "ButtonA_03": outBoolean = False
    Case Is = "ButtonA_04": outBoolean = False
    Case Is = "ButtonA_05": outBoolean = False
    
    Case Is = "PriceApprovalGroupB": outBoolean = True
    Case Is = "ButtonB_01": outBoolean = True
    Case Is = "ButtonB_02": outBoolean = False
    Case Is = "ButtonB_03": outBoolean = False
    Case Is = "ButtonB_04": outBoolean = False
    Case Is = "ButtonB_05": outBoolean = False
    
    Case Is = "PriceApprovalGroupC": outBoolean = True
    Case Is = "LabelC_01": outBoolean = True
    Case Is = "ButtonC_01": outBoolean = True
    
    Case Else: outBoolean = False
    
    End Select
    
    IRibbon_IsVisible = outBoolean
    
End Function


