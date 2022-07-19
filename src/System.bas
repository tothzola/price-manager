Attribute VB_Name = "System"
'@Folder("System.Excel")
Option Explicit

#If VBA7 And Win64 Then
    Private Stat As LongPtr
    Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" _
    (lpdwFlags As LongPtr, ByVal dwReserved As Long) As Boolean
#Else
    Private Stat As Long
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" _
                             (lpdwFlags As Long, ByVal dwReserved As Long) As Boolean
#End If

Private Const WEB_USER_GUIDES As String = "https://google.com"
Private Const WEB_ONLINE_FORMS As String = "https://google.com"
Private Const APP_VERSION As String = "PRICEAPPROVAL_APPVERSION"
Private Ribbon As Office.IRibbonUI
Private Update As IUpdate

Public Property Get LoggerName() As String
    LoggerName = VBA.Left$(ThisWorkbook.Name, VBA.InStr(ThisWorkbook.Name, ".") - 1)
End Property


Public Property Get LoggerFile() As String
    LoggerFile = ThisWorkbook.Path & Application.PathSeparator & LoggerName & ".log"
End Property


Public Property Get LoggerEnabeled() As Boolean
    LoggerEnabeled = LogManager.IsEnabled(LogLevel.DebugLevel)
End Property


Public Property Get CurrentVersion() As Long
    CurrentVersion = CLng(ThisWorkbook.Names(APP_VERSION).RefersToRange.Value)
End Property


Public Function ConnectedToNetwork() As Boolean
    ConnectedToNetwork = (InternetGetConnectedState(Stat, 0&) <> 0)
End Function


Public Sub InitUpdatesManager(ByVal Connection As String, ByVal AppName As String)
    Set Update = UpdateManager.Create(Connection, AppName)
    Dim UpdateAvailable As Boolean
    UpdateAvailable = Update.Available()
End Sub


Public Sub RibbonReCreate(Optional ByRef returnObj As Office.IRibbonUI)

    If Ribbon Is Nothing Then
        Set Ribbon = ObjectBackup.GetObject
    End If

    If Not Ribbon Is Nothing Then
        Ribbon.Invalidate
    End If
    
    Set returnObj = Ribbon

End Sub


'@Description "Determine the text to go along with your Tab, Groups, and Buttons"
Public Function RibbonGetLabel(ByVal Identifier As String, ByRef outLabel As Variant) As Variant
Attribute RibbonGetLabel.VB_Description = "Determine the text to go along with your Tab, Groups, and Buttons"

    Select Case Identifier
    
    'Developer
    Case Is = "PriceApproval_Tab": outLabel = "Price Approval"
    Case Is = "PriceApproval_GroupA": outLabel = "Developer"
    Case Is = "ButtonA_01": outLabel = "Show/ Hide Book"
    Case Is = "ButtonA_02": outLabel = "Label"
    Case Is = "ButtonA_03": outLabel = "Label"
    Case Is = "ButtonA_04": outLabel = "Label"
    Case Is = "ButtonA_05": outLabel = "Label"
    
    'Support
    Case Is = "PriceApproval_GroupB": outLabel = "Support"
    Case Is = "ButtonB_01": outLabel = "Send Feedback"
    Case Is = "ButtonB_02": outLabel = "User Guides"
    Case Is = "LabelB_03": outLabel = " App. Version: " & CurrentVersion
    Case Is = "ButtonB_03": outLabel = "Check for Update"
    Case Is = "ButtonB_04": outLabel = "Label"
    Case Is = "ButtonB_05": outLabel = "Label"
    
    'Applications
    Case Is = "PriceApproval_GroupC": outLabel = "Applications"
    Case Is = "ButtonC_01": outLabel = "Price Approval"
    Case Is = "ButtonC_02": outLabel = "Tapes Approval"
    Case Is = "ButtonC_03": outLabel = "Debitor Approval"
    Case Is = "ButtonC_04": outLabel = "Label"
    Case Is = "ButtonC_05": outLabel = "Label"
    
    Case Else: outLabel = "Button Label"
    
    End Select

    RibbonGetLabel = outLabel
    
End Function

'@Description "Tell each button which macro subroutine to run when clicked"
Public Sub RibbonOnAction(ByVal Control As Office.IRibbonControl)
Attribute RibbonOnAction.VB_Description = "Tell each button which macro subroutine to run when clicked"
    
    If Not LogManager.IsEnabled(LogLevel.DebugLevel) Then
        LogManager.Register FileLogger.Create(LoggerName, DebugLevel, LoggerFile)
    End If
    
    If Control.ID Like "Button*" Then

        Select Case Control.ID

        'Developer
        Case Is = "ButtonA_01": If ThisWorkbook.IsAddin Then ThisWorkbook.IsAddin = False Else ThisWorkbook.IsAddin = True
        Case Is = "ButtonA_02": MsgBox "Not Implementd", vbInformation, PriceApprovalSignature
        Case Is = "ButtonA_03": MsgBox "Not Implementd", vbInformation, PriceApprovalSignature
        Case Is = "ButtonA_04": MsgBox "Not Implementd", vbInformation, PriceApprovalSignature
        Case Is = "ButtonA_05": MsgBox "Not Implementd", vbInformation, PriceApprovalSignature
        
        'Support
        Case Is = "ButtonB_01": If Update.NoUpdate Then EmailServices.EmailFeedback
        Case Is = "ButtonB_02": If Update.NoUpdate Then ThisWorkbook.FollowHyperlink Address:=WEB_USER_GUIDES
        Case Is = "ButtonB_03": If Update.NoUpdate Then Update.Download
        Case Is = "ButtonB_04": If Update.NoUpdate Then MsgBox "Not Implementd", vbInformation, PriceApprovalSignature
        Case Is = "ButtonB_05": If Update.NoUpdate Then MsgBox "Not Implementd", vbInformation, PriceApprovalSignature
        
        'Applications
        Case Is = "ButtonC_01": If Update.NoUpdate Then Main.StartPriceApprovalApp
        Case Is = "ButtonC_02": If Update.NoUpdate Then MsgBox "Not Implementd", vbInformation, PriceApprovalSignature
        Case Is = "ButtonC_03": If Update.NoUpdate Then MsgBox "Not Implementd", vbInformation, PriceApprovalSignature
        Case Is = "ButtonC_04": If Update.NoUpdate Then MsgBox "Not Implementd", vbInformation, PriceApprovalSignature
        Case Is = "ButtonC_05": If Update.NoUpdate Then MsgBox "Not Implementd", vbInformation, PriceApprovalSignature
        
        End Select
        
    End If
    
End Sub

'@Description "Determine the Image to go along with your Buttons"
Public Function RibbonGetImage(ByVal Identifier As String, ByRef outImage As Variant) As Variant
Attribute RibbonGetImage.VB_Description = "Determine the Image to go along with your Buttons"
    
    Select Case Identifier
    
    'Developer
    Case Is = "ButtonA_01": outImage = "DeliverableSynchronize"
    Case Is = "ButtonA_02": outImage = "ObjectPictureFill"
    Case Is = "ButtonA_03": outImage = "ObjectPictureFill"
    Case Is = "ButtonA_04": outImage = "ObjectPictureFill"
    Case Is = "ButtonA_05": outImage = "ObjectPictureFill"
    
    'Support
    Case Is = "ButtonB_01": outImage = "GroupOmsSend"
    Case Is = "ButtonB_02": outImage = "JotShareMenuHelp"
    Case Is = "ButtonB_03": outImage = "RefreshWebView"
    Case Is = "ButtonB_04": outImage = "ObjectPictureFill"
    Case Is = "ButtonB_05": outImage = "ObjectPictureFill"
    
    'Applications
    Case Is = "ButtonC_01": outImage = "CalculatedCurrency"
    Case Is = "ButtonC_02": outImage = "WPColorSchemes"
    Case Is = "ButtonC_03": outImage = "ResourceAllViewsGalleryStandard"
    Case Is = "ButtonC_04": outImage = "ObjectPictureFill"
    Case Is = "Buttonc_05": outImage = "ObjectPictureFill"
    
    Case Else: outImage = "ObjectPictureFill"
    
    End Select
    
    RibbonGetImage = outImage
    
End Function

'@Description "Determine if the button size is large or small"
Public Function RibbonGetSize(ByVal Identifier As String, ByRef outSize As Variant) As Variant
Attribute RibbonGetSize.VB_Description = "Determine if the button size is large or small"

    Const NORMAL As Integer = 0
    Const LARGE As Integer = 1
    
    Select Case Identifier
    
    'Developer
    Case Is = "ButtonA_01": outSize = LARGE
    Case Is = "ButtonA_02": outSize = NORMAL
    Case Is = "ButtonA_03": outSize = NORMAL
    Case Is = "ButtonA_04": outSize = NORMAL
    Case Is = "ButtonA_05": outSize = NORMAL
    
    'Support
    Case Is = "ButtonB_01": outSize = LARGE
    Case Is = "ButtonB_02": outSize = LARGE
    Case Is = "ButtonB_03": outSize = NORMAL
    Case Is = "ButtonB_04": outSize = NORMAL
    Case Is = "ButtonB_05": outSize = NORMAL
    
    'Applications
    Case Is = "ButtonC_01": outSize = LARGE
    Case Is = "ButtonC_02": outSize = LARGE
    Case Is = "ButtonC_03": outSize = LARGE
    Case Is = "ButtonC_04": outSize = NORMAL
    Case Is = "ButtonC_05": outSize = NORMAL
    
    Case Else: outSize = NORMAL
    
    End Select
    
    RibbonGetSize = outSize
    
End Function

'@Description "Display a specific macro description when the mouse hovers over a button"
Public Function RibbonGetScreenTip(ByVal Identifier As String, ByRef outTipp As Variant) As Variant
Attribute RibbonGetScreenTip.VB_Description = "Display a specific macro description when the mouse hovers over a button"
    
    Select Case Identifier
    
    'Developer
    Case Is = "ButtonA_01": outTipp = "Show ThisWorkbook to set app version"
    Case Is = "ButtonA_02": outTipp = "Description"
    Case Is = "ButtonA_03": outTipp = "Description"
    Case Is = "ButtonA_04": outTipp = "Description"
    Case Is = "ButtonA_05": outTipp = "Description"
    
    'Support
    Case Is = "ButtonB_01": outTipp = "Send Feedback to Developers"
    Case Is = "ButtonB_02": outTipp = "Open user guides, open FMP Website"
    Case Is = "ButtonB_03": outTipp = "Get new Updated version"
    Case Is = "ButtonB_04": outTipp = "Description"
    Case Is = "ButtonB_05": outTipp = "Description"
    
    'Applications
    Case Is = "ButtonC_01": outTipp = "Price Approval Application"
    Case Is = "ButtonC_02": outTipp = "Go To Tapes Web Application"
    Case Is = "ButtonC_03": outTipp = "Go To Debitor Web Application"
    Case Is = "ButtonC_04": outTipp = "Description"
    Case Is = "ButtonC_05": outTipp = "Description"
    
    Case Else: outTipp = "Description"
    
    End Select
    
    RibbonGetScreenTip = outTipp
    
End Function

'@Description "Show/Hide buttons based on how many you need"
Public Function RibbonIsVisible(ByVal Control As Office.IRibbonControl, ByRef outBoolean As Variant) As Variant
Attribute RibbonIsVisible.VB_Description = "Show/Hide buttons based on how many you need"

    Select Case Control.ID
    
    'Developer
    Case Is = "PriceApproval_Tab": outBoolean = True
    Case Is = "PriceApproval_GroupA": outBoolean = (Application.UserName = DEVELOPER_NAME)
    Case Is = "ButtonA_01": outBoolean = True
    Case Is = "ButtonA_02": outBoolean = False
    Case Is = "ButtonA_03": outBoolean = False
    Case Is = "ButtonA_04": outBoolean = False
    Case Is = "ButtonA_05": outBoolean = False
    
    'Support
    Case Is = "PriceApproval_GroupB": outBoolean = True
    Case Is = "ButtonB_01": outBoolean = True
    Case Is = "ButtonB_02": outBoolean = True
    Case Is = "LabelB_03": outBoolean = True
    Case Is = "ButtonB_03": outBoolean = True
    Case Is = "ButtonB_04": outBoolean = False
    Case Is = "ButtonB_05": outBoolean = False

    'Applications
    Case Is = "PriceApproval_GroupC": outBoolean = True
    Case Is = "ButtonC_01": outBoolean = True
    Case Is = "ButtonC_02": outBoolean = True
    Case Is = "ButtonC_03": outBoolean = True
    Case Is = "ButtonC_04": outBoolean = False
    Case Is = "ButtonC_05": outBoolean = False
    
    Case Else: outBoolean = False
    
    End Select
    
    RibbonIsVisible = outBoolean
    
End Function
