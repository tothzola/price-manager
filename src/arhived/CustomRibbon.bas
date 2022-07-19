Attribute VB_Name = "CustomRibbon"
Attribute VB_Description = "Ribbon callback, captures the ribbon to an Instance class"
'@IgnoreModule ParameterNotUsed, VariableNotUsed
'@Folder("System.Ribbon")
Option Explicit
Option Private Module

#Const LateBind = TestMode

Private Ribbon As IRibbon
Private Const ModuleName As String = "CustomRibbon"

'@EntryPoint
'@ModuleDescription "Ribbon callback, captures the ribbon to an Instance class"
Public Sub MVPPriceApprovalRibbon(ByVal RibbonUI As Office.IRibbonUI)
    
    On Error GoTo CleanFail
    
    If Not ReferenceCheck.CheckReferenceCompatibility Then Exit Sub
    
    Dim App As Excel.Application
    Set App = Excel.Application
    If App.ProtectedViewWindows.Count > 0 Then Exit Sub

    'Load Ribbon
    Set Ribbon = RibbonManager.Create(RibbonUI)
    Dim WeakRibbon As IWeakReference
    Set WeakRibbon = WeakReference.Create(RibbonUI)
    
CleanExit:
    Exit Sub

CleanFail:
    Err.Raise VBA.vbObjectError + 1091&, ModuleName & ".PriceApprovalRibbon", "Ribbon Load Failed"
    Resume CleanExit
    Resume

End Sub

'@EntryPoint
'@Description "Ribbon Invalidate & Recreate"
Public Sub MVPPriceApprovalInvalidateRibbon()
Attribute MVPPriceApprovalInvalidateRibbon.VB_Description = "Ribbon Invalidate & Recreate"

    On Error GoTo CleanFail
    
    If Not ReferenceCheck.CheckReferenceCompatibility Then Exit Sub
    
    If Ribbon Is Nothing Then
        Dim WeakRibbon As Office.IRibbonUI
        Set WeakRibbon = WeakReference.Ribbon
        Set Ribbon = RibbonManager.Create(WeakRibbon)
    End If
    
    If Not Ribbon Is Nothing Then
        Ribbon.Invalidate

    End If

CleanExit:
    Exit Sub

CleanFail:
    Err.Raise VBA.vbObjectError + 1091&, ModuleName & ".MVPPriceApprovalInvalidateRibbon", "Ribbon Invalidate Failed"
    Resume CleanExit
    Resume

End Sub

Private Sub DebugOutput(ByVal message As String)

    Dim DebugToImmediate As Boolean

    #If LateBind Then
        DebugToImmediate = True
    #End If
    
    CustomRibbon.MVPPriceApprovalInvalidateRibbon
    If DebugToImmediate Then Debug.Print message & "Ribbon was invalidated"
    
End Sub

'@Description "Determine the text to go along with your Tab, Groups, and Buttons"
'@EntryPoint
Private Function MVPPriceApprovalGetLabel(ByVal Control As Office.IRibbonControl, ByRef outLabel As Variant) As Variant
Attribute MVPPriceApprovalGetLabel.VB_Description = "Determine the text to go along with your Tab, Groups, and Buttons"

    Dim returnLabel As Variant
    If Not Ribbon Is Nothing Then
        returnLabel = Ribbon.GetLabel(Control.ID, outLabel)
    Else
        DebugOutput "MVP_PriceApprovalGetLabel: " & Control.ID
        returnLabel = "Label"
    End If
    MVPPriceApprovalGetLabel = returnLabel

End Function

'@Description "Tell Button which macro subroutine to run when clicked"
'@EntryPoint
Private Sub MVPPriceApprovalOnAction(ByVal Control As Office.IRibbonControl)
Attribute MVPPriceApprovalOnAction.VB_Description = "Tell Button which macro subroutine to run when clicked"

    If Not Ribbon Is Nothing Then
        Ribbon.OnAction Control
    ElseIf Control.ID = "ButtonA_03" Then
        CustomRibbon.MVPPriceApprovalInvalidateRibbon
    Else
        DebugOutput "Invalid MVP_PriceApprovalOnAction: " & Control.ID
    End If

End Sub

'@Description "Tell each button which image to load from the Microsoft Icon Library"
'@EntryPoint
Private Function MVPPriceApprovalGetImage(ByVal Control As Office.IRibbonControl, ByRef outImage As Variant) As Variant
Attribute MVPPriceApprovalGetImage.VB_Description = "Tell each button which image to load from the Microsoft Icon Library"

    Dim returnImage As Variant
    If Not Ribbon Is Nothing Then
        returnImage = Ribbon.GetImage(Control.ID, outImage)
    Else
        DebugOutput "MVP_PriceApprovalGetImage: " & Control.ID
        returnImage = "ObjectPictureFill"
    End If
    MVPPriceApprovalGetImage = returnImage

End Function

'@Description "Determine if the button size is large or small"
'@EntryPoint
Private Function MVPPriceApprovalGetSize(ByVal Control As Office.IRibbonControl, ByRef outSize As Variant) As Variant
Attribute MVPPriceApprovalGetSize.VB_Description = "Determine if the button size is large or small"

    Const SMALL As Integer = 0
    Dim returnSize As Variant
    If Not Ribbon Is Nothing Then
        returnSize = Ribbon.GetSize(Control.ID, outSize)
    Else
        DebugOutput "MVP_PriceApprovalGetSize: " & Control.ID
        returnSize = SMALL
    End If
    MVPPriceApprovalGetSize = returnSize

End Function

'@Description "Display a specific macro description when the mouse hovers over a button"
'@EntryPoint
Private Function MVPPriceApprovalGetScreentip(ByVal Control As Office.IRibbonControl, ByRef outTipp As Variant) As Variant
Attribute MVPPriceApprovalGetScreentip.VB_Description = "Display a specific macro description when the mouse hovers over a button"

    Dim returnScreentip As Variant
    If Not Ribbon Is Nothing Then
        returnScreentip = Ribbon.GetScreenTip(Control.ID, outTipp)
    Else
        DebugOutput "MVP_PriceApprovalGetScreentip: " & Control.ID
        returnScreentip = "Description"
    End If
    MVPPriceApprovalGetScreentip = returnScreentip

End Function

'@Description "Show/Hide buttons based on how many you need (False = Hide/True = Show)"
'@EntryPoint
Private Function MVPPriceApprovalGetVisible(ByVal Control As Office.IRibbonControl, ByRef outBoolean As Variant) As Variant
Attribute MVPPriceApprovalGetVisible.VB_Description = "Show/Hide buttons based on how many you need (False = Hide/True = Show)"

    Dim returnShow As Variant
    If Not Ribbon Is Nothing Then
        returnShow = Ribbon.IsVisible(Control, outBoolean)
    Else
        DebugOutput "MVP_PriceApprovalGetVisible: " & Control.ID
        returnShow = True
    End If
    MVPPriceApprovalGetVisible = returnShow

End Function

