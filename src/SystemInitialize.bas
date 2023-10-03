Attribute VB_Name = "SystemInitialize"
Attribute VB_Description = "Ribbon callback, captures the ribbon to an Instance class"
'@IgnoreModule ParameterNotUsed, VariableNotUsed
'@Folder("System")
Option Explicit
Option Private Module

Public Const APP_SIGNATURE As String = "Price Approval"

Private Const APP_NAME As String = "PriceApproval"
Private Const APP_ADDIN_NAME As String = APP_NAME & ".xlam"

Private Const SHAREPOINT_CONNECTION As String = "<LIST><VIEWGUID>{6925B711-7BC0-4090-AB33-27D9659EDA5C}</VIEWGUID><LISTNAME>{8754B0B5-5710-4FFA-B002-DFD1099CE3D3}</LISTNAME><LISTWEB>" & _
                                 "https://sharepoint.site.com/_vti_bin</LISTWEB><LISTSUBWEB></LISTSUBWEB>" & _
                                 "<ROOTFOLDER>/sites/CommonAddIns</ROOTFOLDER></LIST>"
                                 
Private Ribbon As Office.IRibbonUI


'@EntryPoint
'@ModuleDescription "Ribbon callback, captures the ribbon to an Instance class"
Public Sub PriceApproval_SystemStartup(ByVal RibbonUI As Office.IRibbonUI)
    
    On Error GoTo CleanFail
    
    If Not ReferenceCheck.SystemCompatibility(APP_ADDIN_NAME) Then Exit Sub
    If Application.ProtectedViewWindows.Count > 0 Then Exit Sub

    Set Ribbon = RibbonUI
    ObjectBackup.AddObject Ribbon
    System.InitUpdatesManager SHAREPOINT_CONNECTION, APP_ADDIN_NAME
    
CleanExit:
    Exit Sub

CleanFail:
    Err.Raise ErrNo.ObjectSetErr, "PriceApproval SystemStartup", "System Startup Failed"
    Resume CleanExit
    Resume

End Sub

'@Description "Get Ribbon from Backup Object"
Private Function SystemReInitialized() As Boolean
Attribute SystemReInitialized.VB_Description = "Get Ribbon from Backup Object"

    On Error GoTo CleanFail
    
    Dim result As Boolean
    System.RibbonReCreate returnObj:=Ribbon
    System.InitUpdatesManager SHAREPOINT_CONNECTION, APP_ADDIN_NAME
    result = (Not Ribbon Is Nothing)
    
CleanExit:
    SystemReInitialized = result
    Exit Function

CleanFail:
    Err.Raise ErrNo.ObjectSetErr, "Get Ribbon from Backup Failed", "System Reinitialize Failed"
    Resume CleanExit
    Resume
    
End Function


'@Description "Determine the text to go along with your Tab, Groups, and Buttons"
Private Function PriceApproval_GetLabel(ByVal Control As Office.IRibbonControl, ByRef outLabel As Variant) As Variant
Attribute PriceApproval_GetLabel.VB_Description = "Determine the text to go along with your Tab, Groups, and Buttons"

    Dim returnLabel As Variant
    If Not Ribbon Is Nothing Then
        returnLabel = System.RibbonGetLabel(Control.ID, outLabel)
    ElseIf SystemReInitialized Then
        returnLabel = System.RibbonGetLabel(Control.ID, outLabel)
    End If
    PriceApproval_GetLabel = returnLabel

End Function

'@Description "Tell Button which macro subroutine to run when clicked"
Private Sub PriceApproval_OnAction(ByVal Control As Office.IRibbonControl)
Attribute PriceApproval_OnAction.VB_Description = "Tell Button which macro subroutine to run when clicked"
    
    If Not Ribbon Is Nothing Then
        System.RibbonOnAction Control
    ElseIf SystemReInitialized Then
        System.RibbonOnAction Control
    End If

End Sub

'@Description "Tell each button which image to load from the Microsoft Icon Library"
Private Function PriceApproval_GetImage(ByVal Control As Office.IRibbonControl, ByRef outImage As Variant) As Variant
Attribute PriceApproval_GetImage.VB_Description = "Tell each button which image to load from the Microsoft Icon Library"

    Dim returnImage As Variant
    If Not Ribbon Is Nothing Then
        returnImage = System.RibbonGetImage(Control.ID, outImage)
    End If
    PriceApproval_GetImage = returnImage

End Function

'@Description "Determine if the button size is large or small"
Private Function PriceApproval_GetSize(ByVal Control As Office.IRibbonControl, ByRef outSize As Variant) As Variant
Attribute PriceApproval_GetSize.VB_Description = "Determine if the button size is large or small"

    Dim returnSize As Variant
    If Not Ribbon Is Nothing Then
        returnSize = System.RibbonGetSize(Control.ID, outSize)
    End If
    PriceApproval_GetSize = returnSize

End Function

'@Description "Display a specific macro description when the mouse hovers over a button"
Private Function PriceApproval_GetScreentip(ByVal Control As Office.IRibbonControl, ByRef outTipp As Variant) As Variant
Attribute PriceApproval_GetScreentip.VB_Description = "Display a specific macro description when the mouse hovers over a button"

    Dim returnScreentip As Variant
    If Not Ribbon Is Nothing Then
        returnScreentip = System.RibbonGetScreenTip(Control.ID, outTipp)
    End If
    PriceApproval_GetScreentip = returnScreentip

End Function

'@Description "Show/Hide buttons based on how many you need (False = Hide/True = Show)"
Private Function PriceApproval_GetVisible(ByVal Control As Office.IRibbonControl, ByRef outBoolean As Variant) As Variant
Attribute PriceApproval_GetVisible.VB_Description = "Show/Hide buttons based on how many you need (False = Hide/True = Show)"

    Dim returnShow As Variant
    If Not Ribbon Is Nothing Then
        returnShow = System.RibbonIsVisible(Control, outBoolean)
    End If
    PriceApproval_GetVisible = returnShow

End Function
