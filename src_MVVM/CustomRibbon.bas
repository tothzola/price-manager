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
Public Sub MVVMPriceApprovalRibbon(ByVal RibbonUI As Office.IRibbonUI)
    
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
    Err.Raise VBA.vbObjectError + 1091&, ModuleName & ".MVVMPriceApprovalRibbon", "Ribbon Load Failed"
    Resume CleanExit
    Resume

End Sub

'@EntryPoint
'@Description "Ribbon Invalidate & Recreate"
Public Sub MVVMPriceApprovalInvalidateRibbon()
Attribute MVVMPriceApprovalInvalidateRibbon.VB_Description = "Ribbon Invalidate & Recreate"

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
    Err.Raise VBA.vbObjectError + 1091&, ModuleName & ".MVVMPriceApprovalInvalidateRibbon", "Ribbon Invalidate Failed"
    Resume CleanExit
    Resume

End Sub

Private Sub DebugOutput(ByVal Message As String)

    Dim DebugToImmediate As Boolean

    #If LateBind Then
        DebugToImmediate = True
    #End If
    
    CustomRibbon.MVVMPriceApprovalInvalidateRibbon
    If DebugToImmediate Then Debug.Print Message & "Ribbon was invalidated"
    
End Sub

'@Description "Determine the text to go along with your Tab, Groups, and Buttons"
'@EntryPoint
Private Function MVVMPriceApprovalGetLabel(ByVal Control As Office.IRibbonControl, ByRef outLabel As Variant) As Variant
Attribute MVVMPriceApprovalGetLabel.VB_Description = "Determine the text to go along with your Tab, Groups, and Buttons"

    Dim returnLabel As Variant
    If Not Ribbon Is Nothing Then
        returnLabel = Ribbon.GetLabel(Control.ID, outLabel)
    Else
        DebugOutput "MVVM_PriceApprovalGetLabel: " & Control.ID
        returnLabel = "Label"
    End If
    MVVMPriceApprovalGetLabel = returnLabel

End Function

'@Description "Tell Button which macro subroutine to run when clicked"
'@EntryPoint
Private Sub MVVMPriceApprovalOnAction(ByVal Control As Office.IRibbonControl)
Attribute MVVMPriceApprovalOnAction.VB_Description = "Tell Button which macro subroutine to run when clicked"

    If Not Ribbon Is Nothing Then
        Ribbon.OnAction Control
    ElseIf Control.ID = "ButtonA_03" Then
        CustomRibbon.MVVMPriceApprovalInvalidateRibbon
    Else
        DebugOutput "Invalid MVVM_PriceApprovalOnAction: " & Control.ID
    End If

End Sub

'@Description "Tell each button which image to load from the Microsoft Icon Library"
'@EntryPoint
Private Function MVVMPriceApprovalGetImage(ByVal Control As Office.IRibbonControl, ByRef outImage As Variant) As Variant
Attribute MVVMPriceApprovalGetImage.VB_Description = "Tell each button which image to load from the Microsoft Icon Library"

    Dim returnImage As Variant
    If Not Ribbon Is Nothing Then
        returnImage = Ribbon.GetImage(Control.ID, outImage)
    Else
        DebugOutput "MVVM_PriceApprovalGetImage: " & Control.ID
        returnImage = "ObjectPictureFill"
    End If
    MVVMPriceApprovalGetImage = returnImage

End Function

'@Description "Determine if the button size is large or small"
'@EntryPoint
Private Function MVVMPriceApprovalGetSize(ByVal Control As Office.IRibbonControl, ByRef outSize As Variant) As Variant
Attribute MVVMPriceApprovalGetSize.VB_Description = "Determine if the button size is large or small"

    Const SMALL As Integer = 0
    Dim returnSize As Variant
    If Not Ribbon Is Nothing Then
        returnSize = Ribbon.GetSize(Control.ID, outSize)
    Else
        DebugOutput "MVVM_PriceApprovalGetSize: " & Control.ID
        returnSize = SMALL
    End If
    MVVMPriceApprovalGetSize = returnSize

End Function

'@Description "Display a specific macro description when the mouse hovers over a button"
'@EntryPoint
Private Function MVVMPriceApprovalGetScreentip(ByVal Control As Office.IRibbonControl, ByRef outTipp As Variant) As Variant
Attribute MVVMPriceApprovalGetScreentip.VB_Description = "Display a specific macro description when the mouse hovers over a button"

    Dim returnScreentip As Variant
    If Not Ribbon Is Nothing Then
        returnScreentip = Ribbon.GetScreenTip(Control.ID, outTipp)
    Else
        DebugOutput "MVVM_PriceApprovalGetScreentip: " & Control.ID
        returnScreentip = "Description"
    End If
    MVVMPriceApprovalGetScreentip = returnScreentip

End Function

'@Description "Show/Hide buttons based on how many you need (False = Hide/True = Show)"
'@EntryPoint
Private Function MVVMPriceApprovalGetVisible(ByVal Control As Office.IRibbonControl, ByRef outBoolean As Variant) As Variant
Attribute MVVMPriceApprovalGetVisible.VB_Description = "Show/Hide buttons based on how many you need (False = Hide/True = Show)"

    Dim returnShow As Variant
    If Not Ribbon Is Nothing Then
        returnShow = Ribbon.IsVisible(Control, outBoolean)
    Else
        DebugOutput "MVVM_PriceApprovalGetVisible: " & Control.ID
        returnShow = True
    End If
    MVVMPriceApprovalGetVisible = returnShow

End Function

