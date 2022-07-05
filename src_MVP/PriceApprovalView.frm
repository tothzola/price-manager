VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PriceApprovalView 
   Caption         =   "Price Approval Demo"
   ClientHeight    =   24510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20520
   OleObjectBlob   =   "PriceApprovalView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PriceApprovalView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("MVP.View")
'@Exposed
Option Explicit

Implements IView
Implements ICancellable
Implements IDisposable

'-------------------------------------------------------------------------
'Public Events
'-------------------------------------------------------------------------

'Main Frame Events
Public Event OpenLoginFrame()
Public Event ExitApp()
Public Event Logout()
'Login Frame Events
Public Event Login()
Public Event CloseLoginFrame()
Public Event ResetPassword(ByVal targetUserName As String, ByVal TargetEmailAddress As String)
'Password Manager Frame Events
Public Event OpenPasswordManagerFrame()
Public Event ChangePassword()
Public Event ClosePasswordManagerFrame()
'User manager Frame Events
Public Event OpenUserManagerFrame()
Public Event CloseUserManagerFrame()
Public Event ResetUserManagerFrame()
Public Event UpdateUserManagerFrameRecord()
Public Event DoCRUDOperationForUserManager(ByVal TypeOfOperation As CRUDOperations)
'Price Form Frame Events
Public Event OpenPriceFormFrame()
Public Event ClosePriceFormFrame()
Public Event ResetPriceFormFrame()
Public Event DoCRUDOperationForPriceForm(ByVal TypeOfOperation As CRUDOperations)
'Data Form Frame Events
Public Event OpenDataFormFrame(ByVal ContainerIdentifier As DataContainer)
Public Event CloseDataFormFrame()
Public Event ResetDataFormFrame(ByVal ContainerIdentifier As DataContainer)
Public Event EditRecordFromDataFormFrame()
Public Event FilterAndSortListFromDataFormFrame()
Public Event PopulateValuesList()
'Export Form Frame Events
Public Event OpenExportFormFrame()
Public Event CloseExportFormFrame()
Public Event ResetExportFormFrame()
Public Event ExportReport()
             
'-------------------------------------------------------------------------
'VIEW SETTINGS
'-------------------------------------------------------------------------

Private Const FORM_DEF_HEIGHT As Double = 480
Private Const FORM_DEF_WIDTH As Double = 600
Private Const FRAME_INFO_HEIGHT As Double = 78
Private Const FRAME_MARGIN_SIDE As Double = 6
Private Const FRAME_MARGIN_DOUBLE As Double = 12
Private Const FRAME_SIDE_WIDTH As Double = 140
Private Const MESSAGE_WELCOMESCREEN_LOGOUT_STATE As String = "Welcome to The Price Approval Manager"
Private Const MESSAGE_WELCOMESCREEN_LOGIN_STATE As String = "Welcome "

'-------------------------------------------------------------------------
'private type
'-------------------------------------------------------------------------

Private Type TView
    IsCancelled As Boolean
    Resizer As IResizeView
    IsDefaultSizeSet As Boolean
    ViewExtended As MultiFrameViewExtended
    MainModel As AppModel
    LoginModel As LoginFormModel
    PasswordModel As PasswordManagerModel
    UserModel As UserManagerModel
    PriceModel As PriceFormModel
    DataModel As DataFormModel
    ExportModel As ExportFormModel
    Calendar As VBA.Collection
    Disposed As Boolean
End Type

Private This As TView

'-------------------------------------------------------------------------
'Private Variables/Objects
'-------------------------------------------------------------------------

Private EventStop As Boolean

'-------------------------------------------------------------------------
'Supervised Model Properties
'-------------------------------------------------------------------------

Public Property Get MainModel() As AppModel
    Guard.DefaultInstance Me
    Set MainModel = This.MainModel
End Property

Public Property Set MainModel(ByVal RHS As AppModel)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization This.MainModel
    Guard.NullReference RHS
    Set This.MainModel = RHS
End Property

Public Property Get ViewExtended() As MultiFrameViewExtended
    Guard.DefaultInstance Me
    Set ViewExtended = This.ViewExtended
End Property

Public Property Set ViewExtended(ByVal RHS As MultiFrameViewExtended)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization This.ViewExtended
    Guard.NullReference RHS
    Set This.ViewExtended = RHS
End Property

Public Property Get Resizer() As IResizeView
    Guard.DefaultInstance Me
    Set Resizer = This.Resizer
End Property

Public Property Set Resizer(ByVal RHS As IResizeView)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization This.Resizer
    Guard.NullReference RHS
    Set This.Resizer = RHS
End Property

Public Property Get IsDefaultSizeSet() As Boolean
    IsDefaultSizeSet = This.IsDefaultSizeSet
End Property

Public Property Let IsDefaultSizeSet(ByVal RHS As Boolean)
    This.IsDefaultSizeSet = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get LoginModel() As LoginFormModel
    Guard.DefaultInstance Me
    Set LoginModel = This.LoginModel
End Property

Private Property Set LoginModel(ByVal RHS As LoginFormModel)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization This.LoginModel
    Guard.NullReference RHS
    Set This.LoginModel = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get PasswordModel() As PasswordManagerModel
    Guard.DefaultInstance Me
    Set PasswordModel = This.PasswordModel
End Property

Private Property Set PasswordModel(ByVal RHS As PasswordManagerModel)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization This.PasswordModel
    Guard.NullReference RHS
    Set This.PasswordModel = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get UserModel() As UserManagerModel
    Guard.DefaultInstance Me
    Set UserModel = This.UserModel
End Property

Private Property Set UserModel(ByVal RHS As UserManagerModel)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization This.UserModel
    Guard.NullReference RHS
    Set This.UserModel = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get PriceModel() As PriceFormModel
    Guard.DefaultInstance Me
    Set PriceModel = This.PriceModel
End Property

Private Property Set PriceModel(ByVal RHS As PriceFormModel)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization This.PriceModel
    Guard.NullReference RHS
    Set This.PriceModel = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get DataModel() As DataFormModel
    Guard.DefaultInstance Me
    Set DataModel = This.DataModel
End Property

Private Property Set DataModel(ByVal RHS As DataFormModel)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization This.DataModel
    Guard.NullReference RHS
    Set This.DataModel = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get ExportModel() As ExportFormModel
    Guard.DefaultInstance Me
    Set ExportModel = This.ExportModel
End Property

Private Property Set ExportModel(ByVal RHS As ExportFormModel)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization This.ExportModel
    Guard.NullReference RHS
    Set This.ExportModel = RHS
End Property

'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As PriceApprovalView
Attribute Class.VB_Description = "Returns class reference"
    Set Class = PriceApprovalView
End Property

'@Description "Creates a new instance of this form."
Public Function Create(ByVal Model As AppModel) As IView
    
    Guard.NonDefaultInstance Me
    Guard.NullReference Model
    
    Dim result As PriceApprovalView
    Set result = New PriceApprovalView
    
    Set result.MainModel = Model
    Set result.ViewExtended = New MultiFrameViewExtended
    Set result.Resizer = ResizeView.Create(result, FORM_DEF_HEIGHT, FORM_DEF_WIDTH)
    
        InitilizeViewExtended result

    Set Create = result

End Function

'-------------------------------------------------------------------------
'InIt View Method
'-------------------------------------------------------------------------
Private Sub InitilizeViewExtended(ByVal context As PriceApprovalView)
    
    With context.ViewExtended
        'Re-Dimension UserForm
        Set .TargetForm = context
        .formWidth = FORM_DEF_WIDTH
        .formHeight = FORM_DEF_HEIGHT
        .ReDimensionForm
        'Always On Frame Properties
        Set .frameAlwaysOn = context.frameInfo
        .alwaysOnTop = FRAME_MARGIN_SIDE
        .alwaysOnLeft = FRAME_MARGIN_SIDE
        .alwaysOnWidth = FRAME_SIDE_WIDTH
        .alwaysOnHeight = FRAME_INFO_HEIGHT
        'Side Panel Frames Properties
        .sideFrameTop = FRAME_INFO_HEIGHT + FRAME_MARGIN_DOUBLE
        .sideFrameLeft = FRAME_MARGIN_SIDE
        .sideFrameWidth = FRAME_SIDE_WIDTH
        .sideFrameHeight = context.InsideHeight - .alwaysOnHeight - FRAME_MARGIN_DOUBLE - FRAME_MARGIN_SIDE
        'Main Panel Frames Properties
        .mainFrameTop = FRAME_MARGIN_SIDE
        .mainFrameLeft = FRAME_SIDE_WIDTH + FRAME_MARGIN_DOUBLE
        .mainFrameWidth = context.InsideWidth - .sideFrameWidth - FRAME_MARGIN_DOUBLE - FRAME_MARGIN_SIDE
        .mainFrameHeight = context.InsideHeight - (FRAME_MARGIN_DOUBLE)
        'plug static data sources to the relative comboboxes
        .HydrateComboBox context.cmbCurrency, DataResources.arrListofCurrencies
        .HydrateComboBox context.cmbUnitOfMeasure, DataResources.arrListOfUnitOfMeasure
        
        'InIt Interface
        .ActivateFrames context.frameLogin, context.frameWelcome
        .SetDefaultFrameSize context.frameWelcome, "MAIN"
        .SetDefaultFrameSize context.frameLogin, "SIDE"
        .SetDefaultFrameSize context.frameClient, "SIDE"
        .SetDefaultFrameSize context.frameApprover, "SIDE"
        .SetDefaultFrameSize context.frameLoginInterface, "MAIN"
        .SetDefaultFrameSize context.framePasswordManager, "MAIN"
        .SetDefaultFrameSize context.framePriceForm, "MAIN"
        .SetDefaultFrameSize context.frameRecordsContainer, "MAIN"
        .SetDefaultFrameSize context.frameExportUtility, "MAIN"
        .SetDefaultFrameSize context.frameUserManager, "MAIN"
        
        UpdateWelcomeFrame FORM_LOGIN
    End With
    
    'Intit DatePicker
    Set This.Calendar = New VBA.Collection
    Dim i As Integer
    For i = 1 To 42
        This.Calendar.Add New DatePickerFunctions, "titel" & i
        '@Ignore DefaultMemberRequired
        Set This.Calendar("titel" & i).LabelBackground = context("dpLabel" & i)
        If i < 8 Then
            '@Ignore DefaultMemberRequired
            context("dpLabel5" & i).Caption = VBA.Left$(VBA.WeekdayName(i, True, 2), 1)
        End If
    Next

    Dim defaultSizeSet As Boolean
    
    defaultSizeSet = FormControl.MakeFormResizable(context, True)
    defaultSizeSet = FormControl.ShowMinimizeButton(context, False)
    defaultSizeSet = FormControl.ShowMaximizeButton(context, False)
    
    context.IsDefaultSizeSet = defaultSizeSet
    
End Sub

Private Sub MonthsSelector_Change()
    Dim InitDate As Date
    InitDate = VBA.DateSerial(VBA.Year(VBA.Date), VBA.Month(VBA.Date) + Me.MonthsSelector.Value, 1)
    Me.dpLabel50.Caption = VBA.Space(3) & VBA.Year(InitDate) & VBA.Space(6) & VBA.MonthName(VBA.Month(InitDate))

    Dim j As Integer
    For j = 0 To 41
        '@Ignore DefaultMemberRequired
        Me("dpLabel" & j + 1).Caption = VBA.Day(InitDate - VBA.Weekday(InitDate, 2) + 1 + j)
        '@Ignore DefaultMemberRequired
        Me("dpLabel" & j + 1).ForeColor = VBA.IIf(Month(InitDate) = VBA.Month(InitDate - VBA.Weekday(InitDate, 2) + 1 + j), &H80000012, &H80000010)
    Next
End Sub

Private Sub BindControlLayout()

    With Resizer
        .BindControlLayout Me.frameInfo, TopAnchor
        .BindControlLayout Me.frameWelcome, AnchorAll
        .BindControlLayout Me.lblWelcomeMessage, TopAnchor + LeftAnchor + RightAnchor
        .BindControlLayout Me.frameLogin, TopAnchor + BottomAnchor
        .BindControlLayout Me.cmdOpenLoginInterface, TopAnchor
        .BindControlLayout Me.cmdExit, TopAnchor
        
        .BindControlLayout Me.frameClient, TopAnchor + BottomAnchor
        .BindControlLayout Me.cmdOpenPriceForm, TopAnchor
        .BindControlLayout Me.cmdOpenClientHistory, TopAnchor
        .BindControlLayout Me.cmdOpenPasswordManager, TopAnchor
        .BindControlLayout Me.cmdClientLogout, TopAnchor
        
        .BindControlLayout Me.frameApprover, TopAnchor + BottomAnchor
        .BindControlLayout Me.cmdOpenPendingList, TopAnchor
        .BindControlLayout Me.cmdOpenAllHistory, TopAnchor
        .BindControlLayout Me.cmdOpenPasswordManager2, TopAnchor
        .BindControlLayout Me.cmdOpenExportUtility, TopAnchor
        .BindControlLayout Me.cmdOpenUserManager, TopAnchor
        .BindControlLayout Me.cmdApproverLogout, TopAnchor
        
        .BindControlLayout Me.frameLoginInterface, AnchorAll
        .BindControlLayout Me.LoginInterfaceTopPanel, LeftAnchor + RightAnchor
        .BindControlLayout Me.cmdCancelFromLoginInterface, RightAnchor
        
        .BindControlLayout Me.framePasswordManager, AnchorAll
        .BindControlLayout Me.PasswordManagerTopPanel, LeftAnchor + RightAnchor
        .BindControlLayout Me.cmdCancelPasswordManager, RightAnchor

        .BindControlLayout Me.framePriceForm, AnchorAll
        .BindControlLayout Me.PriceFormTopPanel, LeftAnchor + RightAnchor
        .BindControlLayout Me.cmdAddNewRecord, LeftAnchor
        .BindControlLayout Me.cmdUpdateRecord, LeftAnchor
        .BindControlLayout Me.cmdDeleteRecord, LeftAnchor
        .BindControlLayout Me.cmdResetPriceForm, RightAnchor
        .BindControlLayout Me.cmdCancelPriceFormInterface, RightAnchor
        
        .BindControlLayout Me.frameRecordsContainer, AnchorAll
        .BindControlLayout Me.RecordsContainerTopPanel, LeftAnchor + RightAnchor
        .BindControlLayout Me.cmdEditRecord, LeftAnchor
        .BindControlLayout Me.cmdResetDataForm, RightAnchor
        .BindControlLayout Me.cmdCancelRecordContainer, RightAnchor
        .BindControlLayout Me.LabelRecordToEdit, RightAnchor
        .BindControlLayout Me.lstRecordsContainer, AnchorAll
        
        .BindControlLayout Me.frameExportUtility, AnchorAll
        .BindControlLayout Me.ExportUtilityTopPanel, LeftAnchor + RightAnchor
        .BindControlLayout Me.cmdResetExportForm, RightAnchor
        .BindControlLayout Me.cmdCancelExportUtility, RightAnchor
        .BindControlLayout Me.lblMessage, LeftAnchor + RightAnchor
        .BindControlLayout Me.LabelNotes, LeftAnchor + RightAnchor + BottomAnchor
        
        .BindControlLayout Me.frameUserManager, AnchorAll
        .BindControlLayout Me.UserManagerTopPanel, LeftAnchor + RightAnchor
        .BindControlLayout Me.cmdResetUserManager, RightAnchor
        .BindControlLayout Me.cmdCancelUserManager, RightAnchor
        .BindControlLayout Me.LabelItemToUpdate, RightAnchor
        .BindControlLayout Me.lstUsers, AnchorAll
    End With

End Sub

Private Sub InitializeResize()
    If MainModel Is Nothing Then Exit Sub
    BindControlLayout
    Resizer.SetDefaultSize Me
End Sub

Private Sub RedrawView()
    
    Dim viewMinimized As Boolean
    If Not Resizer.IsViewResizable(Me, viewMinimized) Then
        If Not viewMinimized Then
            MsgBox "Minimum View Size Reached." & VBA.vbNewLine & "This is the default minimum size.", _
                vbInformation, SIGN
        End If
    End If

End Sub

Private Sub PozitionCalendar(ByVal Ancor As Variant, ByVal Parent As Variant)
    Me.MonthsSelector.Value = 0
    Me.DatePicker.ZOrder 0
    With Me.DatePicker
        .Left = Ancor.Left + Ancor.Width + Parent.Left + 6
        .Top = Ancor.Top + 6
        .Tag = Ancor.Name
    End With
End Sub

'-------------------------------------------------------------------------
'User Form Events
'-------------------------------------------------------------------------

'Side panel

Private Sub cmdApproverLogout_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent Logout
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdClientLogout_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent Logout
    Me.MousePointer = fmMousePointerDefault
End Sub

'login interface!

Private Sub cmdCancelFromLoginInterface_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent CloseLoginFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdExit_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent ExitApp
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdLogin_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent Login
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdOpenLoginInterface_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent OpenLoginFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub lblForgotPassword_Click()

    Dim targetUserName As String
    If ValidationServices.IsInputValid(outText:=targetUserName, inPrompt:="Please enter UserName: ", inTitel:="Reset Password") Then
        
        Dim targetEmail As String
        If ValidationServices.IsInputValid(outText:=targetEmail, inPrompt:="Please enter registered Email: ", inTitel:="Reset Password") Then
        
            Me.MousePointer = fmMousePointerAppStarting
            VBA.DoEvents
            RaiseEvent ResetPassword(targetUserName, targetEmail)
            Me.MousePointer = fmMousePointerDefault
        
        End If
        
    End If
 
End Sub

'Password Manager
Private Sub cmdOpenPasswordManager_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent OpenPasswordManagerFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdOpenPasswordManager2_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent OpenPasswordManagerFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdCancelPasswordManager_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent ClosePasswordManagerFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdUpdatePassword_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent ChangePassword
    Me.MousePointer = fmMousePointerDefault
End Sub

'User Manager

Private Sub cmdOpenUserManager_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent OpenUserManagerFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdResetUserManager_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent ResetUserManagerFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdCancelUserManager_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent CloseUserManagerFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdAddNewUser_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent DoCRUDOperationForUserManager(CRUD_OPERATION_ADDNEW)
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdDeleteUser_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent DoCRUDOperationForUserManager(CRUD_OPERATION_DELETE)
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdUpdateUser_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent DoCRUDOperationForUserManager(CRUD_OPERATION_UPDATE)
    Me.MousePointer = fmMousePointerDefault
End Sub

'Price Form

Private Sub cmdOpenPriceForm_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent OpenPriceFormFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdResetPriceForm_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent ResetPriceFormFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdCancelPriceFormInterface_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent ClosePriceFormFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdAddNewRecord_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent DoCRUDOperationForPriceForm(CRUD_OPERATION_ADDNEW)
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdUpdateRecord_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    'Hydrate Model Property
    With PriceModel
        .recordStatus = RECORDSTATUS_PENDING
        .statusChangeDate = VBA.Format$(VBA.Now, DATEFORMAT_BACKEND)
    End With
    RaiseEvent DoCRUDOperationForPriceForm(CRUD_OPERATION_UPDATE)
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdApproveRecord_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    'Hydrate Model Property
    With PriceModel
        .recordStatus = RECORDSTATUS_APPROVED
        .statusChangeDate = VBA.Format$(VBA.Now, DATEFORMAT_BACKEND)
    End With
    RaiseEvent DoCRUDOperationForPriceForm(CRUD_OPERATION_APPROVE)
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdRejectRecord_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    'Hydrate Model Property
    With PriceModel
        .recordStatus = RECORDSTATUS_REJECTED
        .statusChangeDate = VBA.Format$(VBA.Now, DATEFORMAT_BACKEND)
    End With
    RaiseEvent DoCRUDOperationForPriceForm(CRUD_OPERATION_REJECT)
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdDeleteRecord_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent DoCRUDOperationForPriceForm(CRUD_OPERATION_DELETE)
    Me.MousePointer = fmMousePointerDefault
End Sub

'Data Form Frame Events

Private Sub cmdOpenAllHistory_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent OpenDataFormFrame(FOR_ALLHISTORY)
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdOpenClientHistory_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent OpenDataFormFrame(FOR_CLIENTHISTORY)
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdOpenPendingList_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent OpenDataFormFrame(FOR_PENDINGAPPROVALS)
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdEditRecord_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent EditRecordFromDataFormFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdCancelRecordContainer_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent CloseDataFormFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

'Export Utility Frame Events

Private Sub cmdOpenExportUtility_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent OpenExportFormFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdCancelExportUtility_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent CloseExportFormFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdResetExportForm_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent ResetExportFormFrame
    Me.MousePointer = fmMousePointerDefault
End Sub

Private Sub cmdExport_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent ExportReport
    Me.MousePointer = fmMousePointerDefault
End Sub

'------------------------------------------------------------------------------
'Export Form Fields Change Events
'------------------------------------------------------------------------------

Private Sub txtDateFrom_Change()
    'Hydrate model property
    ExportModel.FromDate = Me.txtDateFrom.Text
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation Me.txtDateFrom, _
                                                   ExportModel.IsValidField(ExportFormFields.FIELD_FROMDATE), _
                                                   TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, _
                                                   "Date format must be [" & GetDateFormat & "] and Date should be between " & _
                                                   VBA.Format$(START_OF_THE_CENTURY, GetDateFormat) & " and " & _
                                                   VBA.Format$(VBA.Now, GetDateFormat)
End Sub

Private Sub lblGeneratePassword_Click()
    Dim dummyPassword As String
    dummyPassword = AppMethods.RandomString(10)
    Me.txtSetPassword = dummyPassword
End Sub

Private Sub txtDateFrom_Enter()
    PozitionCalendar Me.txtDateFrom, Me.frameExportUtility
    Me.DatePicker.Visible = True
    Me.txtDateFrom.Locked = True
End Sub

Private Sub txtDateFrom_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    txtDateFrom_Enter
End Sub

Private Sub txtDateFrom_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.DatePicker.Visible = False
    Me.txtDateFrom.Locked = False
End Sub

Private Sub txtDateTo_Change()
    'Hydrate model property
    ExportModel.ToDate = Me.txtDateTo.Text
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation Me.txtDateTo, _
                                                   ExportModel.IsValidField(ExportFormFields.FIELD_TODATE), _
                                                   TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, _
                                                   "Date format must be [" & GetDateFormat & "] and Date should be between " & _
                                                   VBA.Format$(START_OF_THE_CENTURY, GetDateFormat) & " and " & _
                                                   VBA.Format$(VBA.Now, GetDateFormat)
End Sub

Private Sub txtDateTo_Enter()
    PozitionCalendar Me.txtDateTo, Me.frameExportUtility
    Me.DatePicker.Visible = True
    Me.txtDateTo.Locked = True
End Sub

Private Sub txtDateTo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    txtDateTo_Enter
End Sub

Private Sub txtDateTo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.DatePicker.Visible = False
    Me.txtDateTo.Locked = False
End Sub

Private Sub cmbCustomerID_Change()
    'hydrate model property
    ExportModel.customerID = Me.cmbCustomerID.Text
    'Validate field
    This.ViewExtended.UpdateControlAfterValidation Me.cmbCustomerID, ExportModel.IsValidField(ExportFormFields.FIELD_CUSTOMERID), TYPE_NA
End Sub

Private Sub cmbUserID_Change()
    'hydrate model property
    ExportModel.userID = Me.cmbUserID.Text
    'Validate field
    This.ViewExtended.UpdateControlAfterValidation Me.cmbUserID, ExportModel.IsValidField(ExportFormFields.FIELD_USERID), TYPE_NA
End Sub

Private Sub cmbStatus_Change()
    'hydrate model property
    ExportModel.recordStatus = Me.cmbStatus.Text
    'Validate field
    This.ViewExtended.UpdateControlAfterValidation Me.cmbStatus, ExportModel.IsValidField(ExportFormFields.FIELD_RECORDSTATUS), TYPE_NA
End Sub

'------------------------------------------------------------------------------
'Data Form Fields Change Events
'------------------------------------------------------------------------------

Private Sub lstRecordsContainer_Click()
    With Me.lstRecordsContainer
        'Hydrate model property
        If .ListIndex > 0 Then
            DataModel.Index = .List(.ListIndex, 0)
            If .List(.ListIndex, 0) = Empty Then
                Me.cmdEditRecord.Enabled = False
            Else
                Me.cmdEditRecord.Enabled = True
            End If
        Else
            Me.cmdEditRecord.Enabled = False
        End If
    End With
End Sub

Private Sub cmdFilterAndSort_Click()
    With Me
        DataModel.selectedColumn = .cmbColumns.Value
        DataModel.selectedValue = .cmbValues.Value
    End With
    RaiseEvent FilterAndSortListFromDataFormFrame
End Sub

Private Sub cmbColumns_Change()
    If Me.cmbColumns.ListIndex > 0 Then
        'Reset Values Combobox Because Columns Combobox has been changed!
        This.ViewExtended.SetStateofControlsToNullState Me.cmbValues
        'Rehydrate Properties
        DataModel.selectedColumn = Me.cmbColumns.Value
        DataModel.selectedValue = Me.cmbValues.Value
        'Raise Event
        RaiseEvent PopulateValuesList
    End If
End Sub

Private Sub cmdResetDataForm_Click()
    Me.MousePointer = fmMousePointerAppStarting
    VBA.DoEvents
    RaiseEvent ResetDataFormFrame(DataModel.ActiveDataContainer)
    Me.MousePointer = fmMousePointerDefault
End Sub

'-------------------------------------------------------------------------
'Price Form Fields Change Events
'-------------------------------------------------------------------------
    
Private Sub txtConditionType_Change()
    'Hydrate Model Property
    PriceModel.conditionType = Me.txtConditionType.Value
    'Validate field
    This.ViewExtended.UpdateControlAfterValidation _
        Me.txtConditionType, _
        PriceModel.IsValidField(MainTableFields.COL_MAIN_ConditionType), _
        TYPE_FIXEDLENGTHSTRING, 4
End Sub

Private Sub cmbSalesOrganization_Change()
    'Hydrate model property
    PriceModel.salesOrganization = Me.cmbSalesOrganization.Value
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation Me.cmbSalesOrganization, _
                                                   PriceModel.IsValidField(MainTableFields.COL_MAIN_SalesOrganization), _
                                                   TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, _
                                                   "This is required field! Please select one option!"
End Sub

Private Sub cmbDistributionChannel_Change()
    'Hydrate model property
    PriceModel.distributionChannel = Me.cmbDistributionChannel.Value
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation _
        Me.cmbDistributionChannel, _
        PriceModel.IsValidField(MainTableFields.COL_Main_DistributionChannel), _
        TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, _
        "This is required field! Please select one option!"
End Sub

Private Sub txtCustomerID_Change()
    'Hydrate model property
    PriceModel.customerID = Me.txtCustomerID.Value
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation _
        Me.txtCustomerID, _
        PriceModel.IsValidField(MainTableFields.COL_MAIN_customerID), _
        TYPE_CUSTOM, _
        "Need exact 6 char length, range should be between [399999] and [599999]"
End Sub

Private Sub txtMaterialID_Change()
    'Hydrate model property
    PriceModel.materialID = Me.txtMaterialID.Value
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation _
        Me.txtMaterialID, _
        PriceModel.IsValidField(MainTableFields.COL_MAIN_materialID), _
        TYPE_CUSTOM, _
        "Need exact 8 char length, range should be between [49999999] and [59999999]"
End Sub

Private Sub txtPrice_Change()
    If EventStop = False Then
        'Event handle mechanism
        EventStop = True
        'Apply formatting
        Me.txtPrice.Value = This.ViewExtended.ApplyFormat(Me.txtPrice.Text, TYPE_CURRENCY)
        'Hydrate model property
        PriceModel.price = Me.txtPrice.Value
        'Validate Field
        This.ViewExtended.UpdateControlAfterValidation _
        Me.txtPrice, _
        PriceModel.IsValidField(MainTableFields.COL_MAIN_price), _
        TYPE_CUSTOM, _
        "maximum 6 char length allowed including decimals!"
        'Event Handle mechanism
        EventStop = False
    End If
End Sub

Private Sub cmbCurrency_Change()
    'Hydrate model property
    PriceModel.currencyType = Me.cmbCurrency.Value
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation _
        Me.cmbCurrency, _
        PriceModel.IsValidField(MainTableFields.COL_MAIN_currency), _
        TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, _
        "This is required field! Please select one option!"
End Sub

Private Sub cmbUnitOfMeasure_Change()
    'Hydrate model property
    PriceModel.unitOfMeasure = Me.cmbUnitOfMeasure.Value
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation _
        Me.cmbUnitOfMeasure, _
        PriceModel.IsValidField(MainTableFields.COL_MAIN_unitOfMeasure), _
        TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, _
        "This is required field! Please select one option!"
End Sub

Private Sub txtPriceUnit_Change()
    'Hydrate model property
    PriceModel.unitOfPrice = Me.txtPriceUnit.Value
    'validate field
    This.ViewExtended.UpdateControlAfterValidation _
        Me.txtPriceUnit, _
        PriceModel.IsValidField(MainTableFields.COL_MAIN_unitOfPrice), _
        TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, _
        "maximal 4 numerical char length"
End Sub

Private Sub txtValidFrom_Change()
    'Hydrate model property
    PriceModel.validFromDate = Me.txtValidFrom.Value
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation _
        Me.txtValidFrom, _
        PriceModel.IsValidField(MainTableFields.COL_MAIN_validFromDate), _
        TYPE_CUSTOM, _
        "Date format must be [" & GetDateFormat & "] and it should be today's date only!"
End Sub

Private Sub txtValidFrom_Enter()
    PozitionCalendar Me.txtValidFrom, Me.framePriceForm
    Me.DatePicker.Visible = True
    Me.txtValidFrom.Locked = True
End Sub

Private Sub txtValidFrom_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    txtValidFrom_Enter
End Sub

Private Sub txtValidFrom_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.DatePicker.Visible = False
    Me.txtValidFrom.Locked = False
End Sub

Private Sub txtValidTo_Change()
    'Hydrate model property
    PriceModel.validToDate = Me.txtValidTo.Value
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation _
        Me.txtValidTo, _
        PriceModel.IsValidField(MainTableFields.COL_MAIN_validToDate), _
        TYPE_CUSTOM, _
        "Date format must be [" & GetDateFormat & "] and it should be future date!"
End Sub

Private Sub txtValidTo_Enter()
    PozitionCalendar Me.txtValidTo, Me.framePriceForm
    Me.DatePicker.Visible = True
    If Not Me.lblActiveUserType.Caption = "APPROVER" Or _
       Not Me.lblActiveUserType.Caption = "MANAGER" Then
        Me.txtValidTo.Locked = True
    End If
End Sub

Private Sub txtValidTo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
    txtValidTo_Enter
End Sub

Private Sub txtValidTo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.DatePicker.Visible = False
    Me.txtValidTo.Locked = False
End Sub

'-------------------------------------------------------------------------
'User Manager Fileds Change Events
'-------------------------------------------------------------------------

Private Sub cmbUserStatus_Change()
    'Hydrate model property
    UserModel.userStatus = Me.cmbUserStatus.Value
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation _
        Me.cmbUserStatus, _
        UserModel.IsValidField(COL_userStatus), _
        TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, _
        "This is required field! Please select one option!"
End Sub

Private Sub cmbUserType_Change()
    'Hydrate model property
    UserModel.userType = Me.cmbUserType.Value
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation Me.cmbUserType, _
                                                   UserModel.IsValidField(COL_userType), _
                                                   TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, _
                                                   "This is required field! Please select one option!"
End Sub

Private Sub txtSetUsername_Change()
    'Hydrate model property
    UserModel.UserName = Me.txtSetUsername.Value
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation Me.txtSetUsername, _
                                                   UserModel.IsValidField(COL_userName), _
                                                   TYPE_CUSTOM, _
                                                   "Username should have minimum 6 characters and it shold be UNIQUE as well."
End Sub

Private Sub txtSetPassword_Change()
    'hydrate model property
    UserModel.userPassword = Me.txtSetPassword.Value
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation Me.txtSetPassword, UserModel.IsValidField(COL_password), TYPE_WRONGPASSWORDPATTERN
End Sub

Private Sub txtUserEmail_Change()
    'hydrate model property
    UserModel.userEmail = Me.txtUserEmail.Value
    'validate field
    This.ViewExtended.UpdateControlAfterValidation Me.txtUserEmail, UserModel.IsValidField(COL_email), TYPE_CUSTOM, "E.g. username@hostname.domain"
End Sub

Private Sub lstUsers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.lstUsers
        If .ListIndex > 0 Then
            If .List(.ListIndex, UsersTableFields.COL_userId - 1) < 102 Then
                'Just for the safetly that they couldn't be able to edit dev's information
                Call This.ViewExtended.ShowMessage("You are not allowed to Update them!", TYPE_INFORMATION)
            Else
                'hydrate model property
                UserModel.userIndex = .List(.ListIndex, 0)
                'Update Record
                RaiseEvent UpdateUserManagerFrameRecord
            End If
        End If
    End With
End Sub

'-------------------------------------------------------------------------
'Password Manager Fields Change Events
'-------------------------------------------------------------------------

Private Sub txtCurrentPassword_Change()
    'Hydrate model property
    PasswordModel.insertedPassword = Me.txtCurrentPassword.Text
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation Me.txtCurrentPassword, PasswordModel.IsValidField(1), TYPE_STRINGSNOTMATCHED
End Sub

Private Sub txtNewPassword_Change()
    'Hydrate model properties
    PasswordModel.NewPassword = Me.txtNewPassword.Text
    'On Every change, of New Password TextBox, We have to reset Confirm Password Field
    Me.txtConfirmNewPassword.Value = vbNullString
    PasswordModel.confirmNewPassword = Me.txtConfirmNewPassword.Text
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation Me.txtNewPassword, PasswordModel.IsValidField(2), TYPE_WRONGPASSWORDPATTERN
End Sub

Private Sub txtConfirmNewPassword_Change()
    'Hydrate model property
    PasswordModel.confirmNewPassword = Me.txtConfirmNewPassword.Text
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation Me.txtConfirmNewPassword, PasswordModel.IsValidField(3), TYPE_STRINGSNOTMATCHED
End Sub

'-------------------------------------------------------------------------
'Login Fields Events
'-------------------------------------------------------------------------

Private Sub txtPassword_Change()
    'Hydrate model property
    LoginModel.Password = Me.txtPassword.Text
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation Me.txtPassword, LoginModel.IsValidPassword, TYPE_NA
End Sub

Private Sub txtUsername_Change()
    'hydrate model property
    LoginModel.UserName = Me.txtUsername.Text
    'Validate Field
    This.ViewExtended.UpdateControlAfterValidation Me.txtUsername, LoginModel.IsValidUsername, TYPE_NA
End Sub

'-------------------------------------------------------------------------
'public Methods Called From Presenters
'-------------------------------------------------------------------------

Public Sub UserWantsToLogout()
    'Logout State
    Call This.ViewExtended.ActivateFrames(Me.frameLogin, Me.frameWelcome)
    Call UpdateWelcomeFrame(FORM_LOGIN)
End Sub

'This Procedure will clode the current frame
Public Sub UserWantsToCloseFrame(ByVal FrameIdentifier As ApplicationForms)
    'open Default Frames
    Select Case FrameIdentifier
    
    Case ApplicationForms.FORM_LOGIN
        Call This.ViewExtended.ActivateFrames(Me.frameLogin, Me.frameWelcome)
        Call UpdateWelcomeFrame(FORM_LOGIN)
            
    Case ApplicationForms.FORM_PASSWORDMANAGER
        Call This.ViewExtended.ActivateFrames(Me.frameClient, Me.frameWelcome)
        Call UpdateWelcomeFrame
            
    Case ApplicationForms.FORM_USERMANAGER
        Call This.ViewExtended.ActivateFrames(Me.frameApprover, Me.frameWelcome)
        Call UpdateWelcomeFrame
            
    Case ApplicationForms.FORM_PRICEFORM
        If MainModel.ActiveUserType = USERTYPE_CLIENT Then
            Call This.ViewExtended.ActivateFrames(Me.frameClient, Me.frameWelcome)
        Else
            Call This.ViewExtended.ActivateFrames(Me.frameApprover, Me.frameWelcome)
        End If
        Call UpdateWelcomeFrame
            
    Case ApplicationForms.FORM_DATAFORM
        If MainModel.ActiveUserType = USERTYPE_CLIENT Then
            Call This.ViewExtended.ActivateFrames(Me.frameClient, Me.frameWelcome)
        Else
            Call This.ViewExtended.ActivateFrames(Me.frameApprover, Me.frameWelcome)
        End If
        Call UpdateWelcomeFrame
            
    Case ApplicationForms.FORM_EXPORTUTILITY
        Call This.ViewExtended.ActivateFrames(Me.frameApprover, Me.frameWelcome)
        Call UpdateWelcomeFrame
            
    End Select
End Sub

Public Sub OnCancel()
    Me.Hide
End Sub

'-------------------------------------------------------------------------
'Public method to open frames
'-------------------------------------------------------------------------

Public Sub UserWantsToOpenLoginFrame(ByVal LoginFrameModel As LoginFormModel)
    'open login interface
    Call This.ViewExtended.ActivateFrames(Me.frameLogin, Me.frameLoginInterface)
    'RESET login frame
    Call ResetLoginFrame(LoginFrameModel)
End Sub

Public Sub UserWantsToOpenPasswordManagerFrame(ByVal PasswordManagerFormModel As PasswordManagerModel)
    'open password manager for the client
    Call This.ViewExtended.ActivateFrames(Me.frameClient, Me.framePasswordManager)
    'RESET Password manager frame
    Call ResetPasswordManagerFrame(PasswordManagerFormModel)
End Sub

Public Sub UserWantsToOpenUserManagerFrame(ByVal UserManagerFormModel As UserManagerModel)
    'open user manager for the client
    Call This.ViewExtended.ActivateFrames(Me.frameApprover, Me.frameUserManager)
    'reset user manager frame
    Call ResetUserManagerFrame(UserManagerFormModel, OPERATION_NEW)
End Sub

Public Sub UserWantsToOpenPriceFormFrame(ByVal PriceFormFrameModel As PriceFormModel, ByVal operation As FormOperation)
    'open Price Form Interface
    If MainModel.ActiveUserType = USERTYPE_CLIENT Then
        Call This.ViewExtended.ActivateFrames(Me.frameClient, Me.framePriceForm)
    Else
        Call This.ViewExtended.ActivateFrames(Me.frameApprover, Me.framePriceForm)
    End If
    'Reset Price Form Frame
    If operation = OPERATION_NEW Then
        Call ResetPriceFormFrame(PriceFormFrameModel, OPERATION_NEW)
    Else
        Call ResetPriceFormFrame(PriceFormFrameModel, OPERATION_UPDATE)
    End If
End Sub

Public Sub UserWantsToOpenExportFormFrame(ByVal ExportFormFrameModel As ExportFormModel)
    'open export form interface
    Call This.ViewExtended.ActivateFrames(Me.frameApprover, Me.frameExportUtility)
    'Reset Export Form Frame
    Call ResetExportFormFrame(ExportFormFrameModel)
End Sub

Public Sub UserWantsToOpenDataFormFrame(ByVal DataFormFrameModel As DataFormModel, ByVal ContainerIdentification As DataContainer)
    'open price form interface
    Select Case ContainerIdentification
    
    Case DataContainer.FOR_CLIENTHISTORY
        'open client history interface
        Call This.ViewExtended.ActivateFrames(Me.frameClient, Me.frameRecordsContainer)
            
    Case DataContainer.FOR_PENDINGAPPROVALS
        'open pending list for approver
        Call This.ViewExtended.ActivateFrames(Me.frameApprover, Me.frameRecordsContainer)
            
    Case DataContainer.FOR_ALLHISTORY
        'open client history interface
        Call This.ViewExtended.ActivateFrames(Me.frameApprover, Me.frameRecordsContainer)
            
    End Select
    'Reset Price Form Frame
    Call ResetDataFormFrame(DataFormFrameModel)
End Sub

'------------------------------------------------------------------------
'Reset Methods
'------------------------------------------------------------------------

Private Sub ResetLoginFrame(ByVal LoginFrameModel As LoginFormModel)
    With Me
        'Attach Model
        If LoginModel Is Nothing Then Set LoginModel = LoginFrameModel
        'clear values of login frame fields
        Call This.ViewExtended.SetStateofControlsToNullState(.txtUsername, .txtPassword)
        'set focus
        .txtUsername.SetFocus
    End With
End Sub

Private Sub ResetPasswordManagerFrame(ByVal PasswordManagerFormModel As PasswordManagerModel)
    With Me
        'Attach Model
        If PasswordModel Is Nothing Then Set PasswordModel = PasswordManagerFormModel
        'clear values of Password manager frame fields
        Call This.ViewExtended.SetStateofControlsToNullState(.txtCurrentPassword, .txtNewPassword, .txtConfirmNewPassword)
        'set focus
        .txtCurrentPassword.SetFocus
    End With
End Sub

Private Sub ResetUserManagerFrame(ByVal UserManagerFormModel As UserManagerModel, ByVal operation As FormOperation)
    With Me
        'Attach Model
        If UserModel Is Nothing Then Set UserModel = UserManagerFormModel
        'clear values of user manager frame fields
        Call This.ViewExtended.SetStateofControlsToNullState(.txtSetUsername, .txtSetPassword, .cmbUserStatus, .cmbUserType, .txtUserEmail, lstUsers)
        'Repopulate ComboBoxes and Listbox
        .cmbUserStatus.List = UserModel.userStatusList
        .cmbUserType.List = UserModel.userTypesList
        With .lstUsers
            .ColumnCount = 7
            .ColumnWidths = "0;0;75;75;75;0;"
            .List = UserModel.usersTable
        End With
        'Put Default Values based on Operation
        If operation = OPERATION_NEW Then
            Call StateForNewRecordForUserManager
            'Set focus
            .txtSetUsername.SetFocus
        Else
            Call StateForUpdateRecordForUserManager
            'Set focus
            .cmbUserStatus.SetFocus
        End If
    End With
End Sub

Private Sub ResetPriceFormFrame(ByVal PriceFormFrameModel As PriceFormModel, ByVal operation As FormOperation)
    With Me
        'Attach Model
        If PriceModel Is Nothing Then Set PriceModel = PriceFormFrameModel
        'clear values of Price form frame fields
        This.ViewExtended.SetStateofControlsToNullState .lblMainRecordStatus, _
                                                        .txtConditionType, _
                                                        .cmbSalesOrganization, _
                                                        .cmbDistributionChannel, _
                                                        .txtCustomerID, _
                                                        .txtMaterialID, _
                                                        .txtPrice, _
                                                        .cmbCurrency, _
                                                        .txtPriceUnit, _
                                                        .cmbUnitOfMeasure, _
                                                        .txtValidFrom, _
                                                        .txtValidTo
        
        'Repopulate ComboBox And ListBox
        .cmbCurrency.List = PriceModel.curenciesList
        .cmbUnitOfMeasure.List = PriceModel.unitOfMeasuresList
        .cmbSalesOrganization.List = PriceModel.salesOrganizationList
        .cmbDistributionChannel.List = PriceModel.distributionChannelList
        'put default values based on operation
        If operation = OPERATION_NEW Then
            Call StateForNewRecordForPriceForm
            'Set Focus
            .cmbDistributionChannel.SetFocus
        ElseIf operation = OPERATION_UPDATE Then
            Call StateForUpdateRecordForPriceForm
        End If
    End With
End Sub

Private Sub ResetDataFormFrame(ByVal DataFormFrameModel As DataFormModel)
    With Me
        'Attach Model
        If DataModel Is Nothing Then Set DataModel = DataFormFrameModel
        'Clear Data Form Controls
        Call This.ViewExtended.SetStateofControlsToNullState(.lstRecordsContainer, .cmbColumns, .cmbValues)
        'Repopulate ListBox and hydrate some of data model properties
        .lblListType = DataModel.ListTitle
        'Filling up listbox with criteria
        With .lstRecordsContainer
            .ColumnCount = 16
            .ColumnWidths = "0;0;;;;0;0;0;;;;;0;0;0;0;"
            .List = DataModel.GetDataForRecordsList
        End With
        .cmbColumns.List = DataModel.DataColumnsList
        'Allow Approver in any case to Approve or Reject Again!
        If MainModel.ActiveUserType = USERTYPE_APPROVER Then
            DataModel.IsApprover = True
        ElseIf MainModel.ActiveUserType = USERTYPE_MANAGER Then
            DataModel.IsManager = True
        Else
            DataModel.IsApprover = False
        End If
        
        'reformat Listbox column with appropriete types
        ReformatListBoxWithAppropriateDataTypesForMainTable
        
        'State of Controls of Data Form
        .cmdEditRecord.Enabled = False
    End With
End Sub

Private Sub ResetExportFormFrame(ByVal ExportFormFrameModel As ExportFormModel)
    With Me
        'Attach Model
        If ExportModel Is Nothing Then Set ExportModel = ExportFormFrameModel
        'Clear Data Form Controls
        Call This.ViewExtended.SetStateofControlsToNullState(.txtDateFrom, .txtDateTo, .cmbCustomerID, .cmbUserID, .cmbStatus, .lblMessage)
        'repopulate comboboxes
        .cmbCustomerID.List = ExportModel.customerIDsList
        .cmbUserID.List = ExportModel.userIDsList
        .cmbStatus.List = ExportModel.statusesList
        'update model
        Call ExportModel.SetPropertiesToDefaultState
        'input field state
        .txtDateFrom.Value = VBA.Format$(ExportModel.FromDate, GetDateFormat)
        .txtDateTo.Value = VBA.Format$(ExportModel.ToDate, GetDateFormat)
        .cmbStatus.Value = ExportModel.recordStatus
    End With
End Sub

'------------------------------------------------------------------------
'Public methods to perfrom operations
'------------------------------------------------------------------------

Public Sub UserWantsToUpdateUserManagerRecord()
    'reset user manager frame
    Call ResetUserManagerFrame(UserModel, OPERATION_UPDATE)
End Sub

Public Sub ShowWarning(ByVal message As String, ByVal typeOfMessage As messageType)
    Call This.ViewExtended.ShowMessage(message, typeOfMessage)
End Sub

Public Sub UserWantsToLogin()
    'Validate Credentials
    Dim response As Variant
    response = LoginModel.IsUserAuthorized
    If response = True Then
        If LoginModel.userStatus = USERSTATUS_ACTIVE Then
            Call OpenNextInterfaceAfterSuccessfulLogin
            Exit Sub
        Else
            Call This.ViewExtended.ShowMessage("Not authorized to LOGIN! Please contact business to know more details.", TYPE_CRITICAL)
        End If
    Else
        Call This.ViewExtended.ShowMessage(response, TYPE_CRITICAL)
    End If
End Sub

Public Sub ApplicationWantsToUpdateValueListComboBox()
    Me.cmbValues.List = DataModel.ValuesList
    Me.cmbValues.SetFocus
End Sub

Public Sub UserWantsToFilterAndSortDataFormList()
    With Me
        'Clear Data Form Controls
        Call This.ViewExtended.SetStateofControlsToNullState(.lstRecordsContainer)
        'Update Listbox
        With .lstRecordsContainer
            .ColumnCount = 16
            .ColumnWidths = "0;0;;;;0;0;0;;;;;0;0;0;0;"
        End With
        If .cmbColumns.Value = vbNullString And .cmbValues.Value = vbNullString Then
            Me.lstRecordsContainer.List = DataModel.GetDataForRecordsList
        Else
            Me.lstRecordsContainer.List = DataModel.GetFilteredAndSortedList
        End If
        'Reformat Grid Columns
        ReformatListBoxWithAppropriateDataTypesForMainTable
    End With
End Sub

'-------------------------------------------------------------------------
'Button Clicked Operations from Main Frame
'-------------------------------------------------------------------------

'Login Frame

Private Sub OpenNextInterfaceAfterSuccessfulLogin()
    'Open Frame based on client type
    If LoginModel.userType = USERTYPE_CLIENT Then
        Call This.ViewExtended.ActivateFrames(Me.frameClient, Me.frameWelcome)
    Else
        Call This.ViewExtended.ActivateFrames(Me.frameApprover, Me.frameWelcome, LoginModel.userType)
    End If
    'Update Active User Frame
    With LoginModel
        Call UpdateActiveUserInfomation(.UserName, .userType, .userStatus, .userID, .Password, .userEmail)
    End With
    'Update Welcome Frame with Username
    Call UpdateWelcomeFrame
End Sub

'Password Manager Frame

Public Sub AfterChangePasswordOperation()
    MsgBox "Password has been changed successfully! Please Sign-In again.", vbInformation, SIGN
    'Go back to logout state
    Call This.ViewExtended.ActivateFrames(Me.frameLogin, Me.frameWelcome)
    Call UpdateWelcomeFrame(FORM_LOGIN)
End Sub

'User Manager Frame

Public Sub AfterUserManagerCRUDOperation(ByVal TypeOfOperation As CRUDOperations, ByVal IsSucceessfullOperation As Boolean)
    Select Case TypeOfOperation
    Case CRUDOperations.CRUD_OPERATION_ADDNEW
        MsgBox "New USER added successfully!", vbInformation, SIGN
    Case CRUDOperations.CRUD_OPERATION_UPDATE
        If IsSucceessfullOperation Then
            MsgBox "User's record has been UPDATED successfully!", vbInformation, SIGN
        End If
    Case CRUDOperations.CRUD_OPERATION_DELETE
        If IsSucceessfullOperation Then
            MsgBox "User has been DELETED successfully!", vbInformation, SIGN
        End If
    End Select
    'Refresh Data Again
    RaiseEvent OpenUserManagerFrame
End Sub

'Price Form Frame
Public Sub AfterPriceFormCRUDOperation(ByVal TypeOfOperation As CRUDOperations, ByVal IsSuccessfullOperation As Boolean)
    Select Case TypeOfOperation
    Case CRUDOperations.CRUD_OPERATION_ADDNEW
        If IsSuccessfullOperation Then
            MsgBox "New Record added successfully!", vbInformation, SIGN
            'reset price form frame
            RaiseEvent OpenPriceFormFrame
        End If
    Case CRUDOperations.CRUD_OPERATION_UPDATE
        If IsSuccessfullOperation Then
            MsgBox "Record has been UPDATED successfully!", vbInformation, SIGN
            'open list for based on user type
            RaiseEvent OpenDataFormFrame(DataModel.ActiveDataContainer)
        End If
    Case CRUDOperations.CRUD_OPERATION_DELETE
        If IsSuccessfullOperation Then
            MsgBox "Record has been DELETED successfully!", vbInformation, SIGN
            'open list for based on user type
            RaiseEvent OpenDataFormFrame(DataModel.ActiveDataContainer)
        End If
    Case CRUDOperations.CRUD_OPERATION_APPROVE
        If IsSuccessfullOperation Then
            MsgBox "Record APPROVED successfully!", vbInformation, SIGN
            'Write here code for sending email to client to notify them
            'open list form based on user type
            RaiseEvent OpenDataFormFrame(DataModel.ActiveDataContainer)
        End If
    Case CRUDOperations.CRUD_OPERATION_REJECT
        If IsSuccessfullOperation Then
            MsgBox "Record REJECTED!", vbInformation, SIGN
            'write here code for sending email to client tp notify them
            'open list form based on user type
            RaiseEvent OpenDataFormFrame(DataModel.ActiveDataContainer)
        End If
    End Select
End Sub

'Export Form Frame

Public Sub AfterExportOperation(ByVal IsSuccessfullOperation As Boolean)
    If IsSuccessfullOperation Then
        RaiseEvent OpenExportFormFrame
    End If
End Sub

'Show Status on Label

Public Sub ShowStatusOfExportProcess(ByVal message As String)
    Me.lblMessage.Caption = message
End Sub

'-------------------------------------------------------------------------
'Methods that helps Reset Procedures!
'-------------------------------------------------------------------------
Private Sub StateForNewRecordForPriceForm()
    With Me
        'update model
        Call PriceModel.SetPropertiesToNewRecordState(MainModel.ActiveUserID)
        'input field state
        .lblMainRecordStatus.Caption = PriceModel.recordStatus
        .txtConditionType.Value = PriceModel.conditionType
        .cmbSalesOrganization.Value = PriceModel.salesOrganization
        .txtPriceUnit.Value = PriceModel.unitOfPrice
        .txtValidFrom.Value = VBA.Format$(PriceModel.validFromDate, GetDateFormat)
        .txtValidTo.Value = VBA.Format$(PriceModel.validToDate, GetDateFormat)
        
        'Hide Buttons
        If MainModel.ActiveUserType = USERTYPE_APPROVER Or MainModel.ActiveUserType = USERTYPE_MANAGER Then
            Call ShowApprovalRejectionButtons(True)
            This.ViewExtended.FormEditingState False, _
                                               .txtConditionType, _
                                               .cmbSalesOrganization, _
                                               .cmbDistributionChannel, _
                                               .txtCustomerID, _
                                               .txtMaterialID, _
                                               .txtPrice, _
                                               .cmbCurrency, _
                                               .txtPriceUnit, _
                                               .cmbUnitOfMeasure
        Else
            Call ShowApprovalRejectionButtons(False)
            This.ViewExtended.FormEditingState True, _
                                               .txtConditionType, _
                                               .cmbSalesOrganization, _
                                               .cmbDistributionChannel, _
                                               .txtCustomerID, _
                                               .txtMaterialID, _
                                               .txtPrice, _
                                               .cmbCurrency, _
                                               .txtPriceUnit, _
                                               .cmbUnitOfMeasure
            
        End If
        'Other Buttons State
        .cmdAddNewRecord.Enabled = True
        .cmdResetPriceForm.Enabled = True
        .cmdUpdateRecord.Enabled = False
        .cmdDeleteRecord.Enabled = False
    End With
End Sub

Private Sub StateForUpdateRecordForPriceForm()
    With Me
        'update model
        Call PriceModel.SetPropertiesToUpdateRecordState
        'input field state
        .lblMainRecordStatus.Caption = PriceModel.recordStatus
        .txtConditionType.Value = PriceModel.conditionType
        .cmbSalesOrganization.Value = PriceModel.salesOrganization
        .cmbDistributionChannel.Value = PriceModel.distributionChannel
        .txtCustomerID.Value = PriceModel.customerID
        .txtMaterialID.Value = PriceModel.materialID
        .txtPrice.Value = PriceModel.price
        .cmbCurrency.Value = PriceModel.currencyType
        .txtPriceUnit.Value = PriceModel.unitOfPrice
        .cmbUnitOfMeasure.Value = PriceModel.unitOfMeasure
        .txtValidFrom.Value = VBA.Format$(PriceModel.validFromDate, GetDateFormat)
        .txtValidTo.Value = VBA.Format$(PriceModel.validToDate, GetDateFormat)
        
        'Hide Buttons & Form Lock Decision
        If MainModel.ActiveUserType = USERTYPE_APPROVER Or MainModel.ActiveUserType = USERTYPE_MANAGER Then
            Call ShowApprovalRejectionButtons(True)
            This.ViewExtended.FormEditingState False, _
                                               .txtConditionType, _
                                               .cmbSalesOrganization, _
                                               .cmbDistributionChannel, _
                                               .txtCustomerID, _
                                               .txtMaterialID, _
                                               .txtPrice, _
                                               .cmbCurrency, _
                                               .txtPriceUnit, _
                                               .cmbUnitOfMeasure
            
            'Other Buttons State
            .cmdAddNewRecord.Enabled = False
            .cmdUpdateRecord.Enabled = False
            .cmdDeleteRecord.Enabled = False
            .cmdResetPriceForm.Enabled = False
        Else
            Call ShowApprovalRejectionButtons(False)
            This.ViewExtended.FormEditingState True, _
                                               .txtConditionType, _
                                               .cmbSalesOrganization, _
                                               .cmbDistributionChannel, _
                                               .txtCustomerID, _
                                               .txtMaterialID, _
                                               .txtPrice, _
                                               .cmbCurrency, _
                                               .txtPriceUnit, _
                                               .cmbUnitOfMeasure
            
            'Other Buttons State
            .cmdAddNewRecord.Enabled = False
            .cmdUpdateRecord.Enabled = True
            .cmdDeleteRecord.Enabled = True
            .cmdResetPriceForm.Enabled = False
        End If
    End With
End Sub

Private Sub StateForNewRecordForUserManager()
    With Me
        'Update Model
        Call UserModel.SetPropertiesToNewUserState
        'Input Field State
        .cmbUserStatus.Value = UserModel.userStatus
        .cmbUserType.Value = UserModel.userType
        'Buttons State
        .cmdAddNewUser.Enabled = True
        .cmdUpdateUser.Enabled = False
        .cmdDeleteUser.Enabled = False
    End With
End Sub

Private Sub StateForUpdateRecordForUserManager()
    With Me
        'Field State
        Call UserModel.SetPropertiesToUpdateUserState
        'input field state
        .cmbUserStatus.Value = UserModel.userStatus
        .cmbUserType.Value = UserModel.userType
        .txtSetUsername.Value = UserModel.UserName
        .txtUserEmail.Value = UserModel.userEmail
        'Button State
        .cmdAddNewUser.Enabled = False
        .cmdUpdateUser.Enabled = True
        .cmdDeleteUser.Enabled = True
    End With
End Sub

Private Sub ShowApprovalRejectionButtons(ByVal decision As Boolean)
    With Me
        .cmdApproveRecord.Visible = decision
        .cmdRejectRecord.Visible = decision
    End With
End Sub

Private Sub UpdateWelcomeFrame(Optional FrameIdentifier As ApplicationForms = 0)
    If FrameIdentifier = FORM_LOGIN Then
        'Update Welcome Frame while user is in logout state
        With MultiFrameViewExtended
            Call .ChangeControlProperties(Me.lblWelcomeMessage, MESSAGE_WELCOMESCREEN_LOGOUT_STATE, &H8000000D)
            Call .SetStateofControlsToNullState(Me.lblActiveUsername, Me.lblActiveUserType, Me.lblActiveUserStatus, Me.lblActiveUserID, Me.lblActiveUserPassword)
        End With
    Else
        'Update Welcome Message While User is Still Logged In
        Dim propperUserName As String
        propperUserName = VBA.Strings.StrConv(Me.lblActiveUsername.Caption, VBA.vbProperCase)
        Call This.ViewExtended.ChangeControlProperties(Me.lblWelcomeMessage, MESSAGE_WELCOMESCREEN_LOGIN_STATE & propperUserName, &H8000000D)
    End If
End Sub

Private Sub UpdateActiveUserInfomation(ByVal uName As String, _
                                       ByVal uType As String, ByVal uStatus As String, ByVal uID As String, ByVal uPassword As String, ByVal uEmail As String)

    'Show Active user info on Always On Frame
    With MultiFrameViewExtended
        Call .ChangeControlProperties(Me.lblActiveUsername, VBA.UCase$(uName))
        Call .ChangeControlProperties(Me.lblActiveUserType, VBA.UCase$(uType))
        Call .ChangeControlProperties(Me.lblActiveUserID, uID)
        Call .ChangeControlProperties(Me.lblActiveUserPassword, uPassword)
        If uStatus = USERSTATUS_ACTIVE Then
            Call .ChangeControlProperties(Me.lblActiveUserStatus, uStatus, COLOR_OF_OKAY)
        Else
            Call .ChangeControlProperties(Me.lblActiveUserStatus, uStatus, COLOR_OF_NOT_OKAY)
        End If
    End With
    'Update Active user information in Main Model
    With MainModel
        .ActiveUserID = uID
        .ActiveUserName = uName
        .ActiveUserPassword = uPassword
        .ActiveUserStatus = uStatus
        .ActiveUserType = uType
        .ActiveUserEmail = uEmail
    End With
End Sub

Private Sub ReformatListBoxWithAppropriateDataTypesForMainTable()
    With Me
        'Edit Change Date
        Call This.ViewExtended.ReformatListBoxColumns(.lstRecordsContainer, MainTableFields.COL_MAIN_statusChangeDate, TYPE_DATE)
        'price column
        Call This.ViewExtended.ReformatListBoxColumns(.lstRecordsContainer, MainTableFields.COL_MAIN_price, TYPE_CURRENCY)
        'From Date Column
        Call This.ViewExtended.ReformatListBoxColumns(.lstRecordsContainer, MainTableFields.COL_MAIN_validFromDate, TYPE_DATE)
        'To Date Column
        Call This.ViewExtended.ReformatListBoxColumns(.lstRecordsContainer, MainTableFields.COL_MAIN_validToDate, TYPE_DATE)
    End With
End Sub

Private Sub Dispose()

    If This.Disposed Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If
    
    If Not This.ViewExtended Is Nothing Then
        Disposable.TryDispose This.ViewExtended
        Set This.ViewExtended = Nothing
    End If
    
    If Not This.MainModel Is Nothing Then
        Disposable.TryDispose This.MainModel
        Set This.MainModel = Nothing
    End If
    
    If Not This.LoginModel Is Nothing Then
        Disposable.TryDispose This.LoginModel
        Set This.LoginModel = Nothing
    End If
    
    If Not This.UserModel Is Nothing Then
        Disposable.TryDispose This.UserModel
        Set This.UserModel = Nothing
    End If
    
    If Not This.PriceModel Is Nothing Then
        Disposable.TryDispose This.PriceModel
        Set This.PriceModel = Nothing
    End If

    If Not This.DataModel Is Nothing Then
        Disposable.TryDispose This.DataModel
        Set This.DataModel = Nothing
    End If
    
    If Not This.ExportModel Is Nothing Then
        Disposable.TryDispose This.ExportModel
        Set This.ExportModel = Nothing
    End If

    If Not This.Resizer Is Nothing Then
        Disposable.TryDispose This.Resizer
        Set This.Resizer = Nothing
    End If
    
    This.Disposed = True
    
    #If TestMode Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " is terminating"
    #End If
    
End Sub

'-------------------------------------------------------------------------
'Userform Events
'-------------------------------------------------------------------------

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub Class_Terminate()
    If Not This.Disposed Then Dispose
End Sub

Private Sub UserForm_Resize()
    If IsDefaultSizeSet Then RedrawView
End Sub

Private Sub IView_Show()
    InitializeResize
    Me.Show vbModeless
End Sub

Private Sub IView_Hide()
    Me.Hide
End Sub

Private Property Get ICancellable_IsCancelled() As Boolean
    ICancellable_IsCancelled = This.IsCancelled
End Property

Private Sub ICancellable_OnCancel()
    OnCancel
End Sub

Private Sub IDisposable_Dispose()
    Dispose
End Sub

