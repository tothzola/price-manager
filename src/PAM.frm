VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PAM 
   Caption         =   "Price Approval Manager V1.0"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2550
   OleObjectBlob   =   "PAM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "View"
Option Explicit

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
Public Event PopulateValuesList(ByVal TargetColumn As String)
'Export Form Frame Events
Public Event OpenExportFormFrame()
Public Event CloseExportFormFrame()
Public Event ResetExportFormFrame()
Public Event ExportReport()
             
'-------------------------------------------------------------------------
'VIEW SETTINGS
'-------------------------------------------------------------------------

Const MESSAGE_WELCOMESCREEN_LOGOUT_STATE As String = "Welcome to The Price Approval Manager"
Const MESSAGE_WELCOMESCREEN_LOGIN_STATE As String = "Welcome "

'-------------------------------------------------------------------------
'private type
'-------------------------------------------------------------------------

Private Type TViewComponents
    MainModel As AppModel
    LoginModel As LoginFormModel
    PasswordModel As PasswordManagerModel
    UserModel As UserManagerModel
    PriceModel As PriceFormModel
    DataModel As DataFormModel
    ExportModel As ExportFormModel
End Type

Private this As TViewComponents

'-------------------------------------------------------------------------
'Private Variables/Objects
'-------------------------------------------------------------------------

Private ExtendedMethods As MultiFrameViewExtended
Private EventStop As Boolean

'-------------------------------------------------------------------------
'Supervised Model Properties
'-------------------------------------------------------------------------

Private Property Get MainModel() As AppModel
    Set MainModel = this.MainModel
End Property

Private Property Set MainModel(ByVal vNewValue As AppModel)
    Set this.MainModel = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get LoginModel() As LoginFormModel
    Set LoginModel = this.LoginModel
End Property

Private Property Set LoginModel(ByVal vNewValue As LoginFormModel)
    Set this.LoginModel = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get PasswordModel() As PasswordManagerModel
    Set PasswordModel = this.PasswordModel
End Property

Private Property Set PasswordModel(ByVal vNewValue As PasswordManagerModel)
    Set this.PasswordModel = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get UserModel() As UserManagerModel
    Set UserModel = this.UserModel
End Property

Private Property Set UserModel(ByVal vNewValue As UserManagerModel)
    Set this.UserModel = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get PriceModel() As PriceFormModel
    Set PriceModel = this.PriceModel
End Property

Private Property Set PriceModel(ByVal vNewValue As PriceFormModel)
    Set this.PriceModel = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get DataModel() As DataFormModel
    Set DataModel = this.DataModel
End Property

Private Property Set DataModel(ByVal vNewValue As DataFormModel)
    Set this.DataModel = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get ExportModel() As ExportFormModel
    Set ExportModel = this.ExportModel
End Property

Private Property Set ExportModel(ByVal vNewValue As ExportFormModel)
    Set this.ExportModel = vNewValue
End Property

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

'Password Manager

Private Sub cmdOpenPasswordManager_Click()
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
    ExtendedMethods.UpdateControlAfterValidation Me.txtDateFrom, ExportModel.IsValidField(ExportFormFields.FIELD_FROMDATE), TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, "Date format must be [DD.MM.YYYY] OR [DDMMYYY] and Date should be between " & VBA.Format$(START_OF_THE_CENTURY, DATEFORMAT_FRONTEND) & " and " & VBA.Format$(VBA.Now, DATEFORMAT_FRONTEND)
End Sub

Private Sub txtDateTo_Change()
    'Hydrate model property
    ExportModel.ToDate = Me.txtDateTo.Text
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtDateTo, ExportModel.IsValidField(ExportFormFields.FIELD_TODATE), TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, "Date format must be [DD.MM.YYYY] OR [DDMMYYY] and Date should be between " & VBA.Format$(START_OF_THE_CENTURY, DATEFORMAT_FRONTEND) & " and " & VBA.Format$(VBA.Now, DATEFORMAT_FRONTEND)
End Sub

Private Sub cmbCustomerID_Change()
    'hydrate model property
    ExportModel.customerID = Me.cmbCustomerID.Text
    'Validate field
    ExtendedMethods.UpdateControlAfterValidation Me.cmbCustomerID, ExportModel.IsValidField(ExportFormFields.FIELD_CUSTOMERID), TYPE_NA
End Sub

Private Sub cmbUserID_Change()
    'hydrate model property
    ExportModel.userID = Me.cmbUserID.Text
    'Validate field
    ExtendedMethods.UpdateControlAfterValidation Me.cmbUserID, ExportModel.IsValidField(ExportFormFields.FIELD_USERID), TYPE_NA
End Sub

Private Sub cmbStatus_Change()
    'hydrate model property
    ExportModel.recordStatus = Me.cmbStatus.Text
    'Validate field
    ExtendedMethods.UpdateControlAfterValidation Me.cmbStatus, ExportModel.IsValidField(ExportFormFields.FIELD_RECORDSTATUS), TYPE_NA
End Sub

'------------------------------------------------------------------------------
'Data Form Fields Change Events
'------------------------------------------------------------------------------

Private Sub lstRecordsContainer_Click()
    With Me.lstRecordsContainer
        'Hydrate model property
        If .ListIndex > 0 Then
            DataModel.index = .List(.ListIndex, 0)
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
        DataModel.selectedColumn = DataModel.GetTargetColumnIndex(.cmbColumns.Value)
        DataModel.selectedValue = .cmbValues.Value
    End With
    RaiseEvent FilterAndSortListFromDataFormFrame
End Sub

Private Sub cmbColumns_Change()
    If Me.cmbColumns.ListIndex > 0 Then
        Me.cmbValues.Value = vbNullString
        RaiseEvent PopulateValuesList(Me.cmbColumns.Value)
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
    ExtendedMethods.UpdateControlAfterValidation Me.txtConditionType, PriceModel.IsValidField(MainTableFields.COL_MAIN_ConditionType), TYPE_FIXEDLENGTHSTRING, 4
End Sub

Private Sub cmbSalesOrganization_Change()
    'Hydrate model property
    PriceModel.salesOrganization = Me.cmbSalesOrganization.Value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.cmbSalesOrganization, PriceModel.IsValidField(MainTableFields.COL_MAIN_SalesOrganization), TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, "This is required field! Please select one option!"
End Sub

Private Sub cmbDistributionChannel_Change()
    'Hydrate model property
    PriceModel.distributionChannel = Me.cmbDistributionChannel.Value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.cmbDistributionChannel, PriceModel.IsValidField(MainTableFields.Col_Main_DistributionChannel), TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, "This is required field! Please select one option!"
End Sub
    
Private Sub txtCustomerID_Change()
    'Hydrate model property
    PriceModel.customerID = Me.txtCustomerID.Value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtCustomerID, PriceModel.IsValidField(MainTableFields.COL_MAIN_customerID), TYPE_CUSTOM, "Need exact 6 char length, range should be between [399999] and [599999]"
End Sub

Private Sub txtMaterialID_Change()
    'Hydrate model property
    PriceModel.materialID = Me.txtMaterialID.Value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtMaterialID, PriceModel.IsValidField(MainTableFields.COL_MAIN_materialID), TYPE_CUSTOM, "Need exact 8 char length, range should be between [49999999] and [59999999]"
End Sub

Private Sub txtPrice_Change()
    If EventStop = False Then
        'Event handle mechanism
        EventStop = True
        'Apply formatting
        Me.txtPrice.Value = ExtendedMethods.ApplyFormat(Me.txtPrice.Text, TYPE_CURRENCY)
        'Hydrate model property
        PriceModel.price = Me.txtPrice.Value
        'Validate Field
        ExtendedMethods.UpdateControlAfterValidation Me.txtPrice, PriceModel.IsValidField(MainTableFields.COL_MAIN_price), TYPE_CUSTOM, "maximum 6 char length allowed including decimals!"
        'Event Handle mechanism
        EventStop = False
    End If
End Sub

Private Sub cmbCurrency_Change()
    'Hydrate model property
    PriceModel.currencyType = Me.cmbCurrency.Value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.cmbCurrency, PriceModel.IsValidField(MainTableFields.COL_MAIN_currency), TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, "This is required field! Please select one option!"
End Sub

Private Sub cmbUnitOfMeasure_Change()
    'Hydrate model property
    PriceModel.unitOfMeasure = Me.cmbUnitOfMeasure.Value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.cmbUnitOfMeasure, PriceModel.IsValidField(MainTableFields.COL_MAIN_unitOfMeasure), TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, "This is required field! Please select one option!"
End Sub

Private Sub txtPriceUnit_Change()
    'Hydrate model property
    PriceModel.unitOfPrice = Me.txtPriceUnit.Value
    'validate field
    ExtendedMethods.UpdateControlAfterValidation Me.txtPriceUnit, PriceModel.IsValidField(MainTableFields.COL_MAIN_unitOfPrice), TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, "maximal 4 numerical char length"
End Sub

Private Sub txtValidFrom_Change()
    'Hydrate model property
    PriceModel.validFromDate = Me.txtValidFrom.Value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtValidFrom, PriceModel.IsValidField(MainTableFields.COL_MAIN_validFromDate), TYPE_CUSTOM, "Date format must be [DD.MM.YYYY] OR [DDMMYYY] and it should be today's date only!"
End Sub

Private Sub txtValidTo_Change()
    'Hydrate model property
    PriceModel.validToDate = Me.txtValidTo.Value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtValidTo, PriceModel.IsValidField(MainTableFields.COL_MAIN_validToDate), TYPE_CUSTOM, "Date format must be [DD.MM.YYYY] OR [DDMMYYY] and it should be future date!"
End Sub
   
'-------------------------------------------------------------------------
'User Manager Fileds Change Events
'-------------------------------------------------------------------------

Private Sub cmbUserStatus_Change()
    'Hydrate model property
    UserModel.userStatus = Me.cmbUserStatus.Value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.cmbUserStatus, UserModel.IsValidField(COL_userStatus), TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, "This is required field! Please select one option!"
End Sub

Private Sub cmbUserType_Change()
    'Hydrate model property
    UserModel.UserType = Me.cmbUserType.Value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.cmbUserType, UserModel.IsValidField(COL_userType), TYPE_AllowBlankButIfValueIsNotNullThenConditionApplied, "This is required field! Please select one option!"
End Sub

Private Sub txtSetUsername_Change()
    'Hydrate model property
    UserModel.UserName = Me.txtSetUsername.Value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtSetUsername, UserModel.IsValidField(COL_userName), TYPE_CUSTOM, "Username should have minimum 6 characters and it shold be UNIQUE as well."
End Sub

Private Sub txtSetPassword_Change()
    'hydrate model property
    UserModel.userPassword = Me.txtSetPassword.Value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtSetPassword, UserModel.IsValidField(COL_password), TYPE_WRONGPASSWORDPATTERN
End Sub

Private Sub txtUserEmail_Change()
    'hydrate model property
    UserModel.userEmail = Me.txtUserEmail.Value
    'validate field
    ExtendedMethods.UpdateControlAfterValidation Me.txtUserEmail, UserModel.IsValidField(COL_email), TYPE_CUSTOM, "E.g. username@hostname.domain"
End Sub

Private Sub lstUsers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.lstUsers
        If .ListIndex > 0 Then
            If .List(.ListIndex, UsersTableFields.COL_userId - 1) = 100 Or _
                .List(.ListIndex, UsersTableFields.COL_userId - 1) = 101 Then
                'Just for the safetly that they couldn't be able to edit dev's information
                Call ExtendedMethods.ShowMessage("You are not allowed to Update them!", TYPE_INFORMATION)
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
    ExtendedMethods.UpdateControlAfterValidation Me.txtCurrentPassword, PasswordModel.IsValidField(1), TYPE_STRINGSNOTMATCHED
End Sub

Private Sub txtNewPassword_Change()
    'Hydrate model properties
    PasswordModel.newPassword = Me.txtNewPassword.Text
    'On Every change, of New Password TextBox, We have to reset Confirm Password Field
    Me.txtConfirmNewPassword.Value = vbNullString
    PasswordModel.confirmNewPassword = Me.txtConfirmNewPassword.Text
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtNewPassword, PasswordModel.IsValidField(2), TYPE_WRONGPASSWORDPATTERN
End Sub

Private Sub txtConfirmNewPassword_Change()
    'Hydrate model property
    PasswordModel.confirmNewPassword = Me.txtConfirmNewPassword.Text
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtConfirmNewPassword, PasswordModel.IsValidField(3), TYPE_STRINGSNOTMATCHED
End Sub

'-------------------------------------------------------------------------
'Login Fields Events
'-------------------------------------------------------------------------

Private Sub txtPassword_Change()
    'Hydrate model property
    LoginModel.Password = Me.txtPassword.Text
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtPassword, LoginModel.IsValidPassword, TYPE_NA
End Sub

Private Sub txtUsername_Change()
    'hydrate model property
    LoginModel.UserName = Me.txtUsername.Text
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtUsername, LoginModel.IsValidUsername, TYPE_NA
End Sub

'-------------------------------------------------------------------------
'InIt View Method
'-------------------------------------------------------------------------

Public Sub InItApplication(ByVal ApplicationModel As AppModel)
    'init Extended Methods
    Set MainModel = ApplicationModel
    Set ExtendedMethods = New MultiFrameViewExtended
    With ExtendedMethods
        'Re-Dimension UserForm
        Set .TargetForm = Me
        .formWidth = 600
        .formHeight = 400
        Call .ReDimensionForm
        'Always On Frame Properties
        Set .frameAlwaysOn = Me.frameInfo
        .alwaysOnTop = 6
        .alwaysOnLeft = 6
        .alwaysOnWidth = 140
        .alwaysOnHeight = 78
        'Side Panel Frames Properties
        .sideFrameTop = 90
        .sideFrameLeft = 6
        .sideFrameWidth = 140
        .sideFrameHeight = 274
        'Main Panel Frames Properties
        .mainFrameTop = 6
        .mainFrameLeft = 152
        .mainFrameWidth = 430
        .mainFrameHeight = 358
        'plug static data sources to the relative comboboxes
        Call .HydrateComboBox(Me.cmbCurrency, modDataSources.arrListofCurrencies)
        Call .HydrateComboBox(Me.cmbUnitOfMeasure, modDataSources.arrListOfUnitOfMeasure)
        'InIt Interface
        Call .ActivateFrames(Me.frameLogin, Me.frameWelcome)
        Call UpdateWelcomeFrame(FORM_LOGIN)
    End With
End Sub

'-------------------------------------------------------------------------
'public Methods Called From Presenters
'-------------------------------------------------------------------------

Public Sub UserWantsToLogout()
    'Logout State
    Call ExtendedMethods.ActivateFrames(Me.frameLogin, Me.frameWelcome)
    Call UpdateWelcomeFrame(FORM_LOGIN)
End Sub

'This Procedure will clode the current frame
Public Sub UserWantsToCloseFrame(ByVal FrameIdentifier As ApplicationForms)
    'open Default Frames
    Select Case FrameIdentifier
    
        Case ApplicationForms.FORM_LOGIN
            Call ExtendedMethods.ActivateFrames(Me.frameLogin, Me.frameWelcome)
            Call UpdateWelcomeFrame(FORM_LOGIN)
            
        Case ApplicationForms.FORM_PASSWORDMANAGER
            Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.frameWelcome)
            Call UpdateWelcomeFrame
            
        Case ApplicationForms.FORM_USERMANAGER
            Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameWelcome)
            Call UpdateWelcomeFrame
            
        Case ApplicationForms.FORM_PRICEFORM
            If MainModel.ActiveUserType = USERTYPE_CLIENT Then
                Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.frameWelcome)
            Else
                Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameWelcome)
            End If
            Call UpdateWelcomeFrame
            
        Case ApplicationForms.FORM_DATAFORM
            If MainModel.ActiveUserType = USERTYPE_CLIENT Then
                Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.frameWelcome)
            Else
                Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameWelcome)
            End If
            Call UpdateWelcomeFrame
            
        Case ApplicationForms.FORM_EXPORTUTILITY
            Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameWelcome)
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
    Call ExtendedMethods.ActivateFrames(Me.frameLogin, Me.frameLoginInterface)
    'RESET login frame
    Call ResetLoginFrame(LoginFrameModel)
End Sub

Public Sub UserWantsToOpenPasswordManagerFrame(ByVal PasswordManagerFormModel As PasswordManagerModel)
    'open password manager for the client
    Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.framePasswordManager)
    'RESET Password manager frame
    Call ResetPasswordManagerFrame(PasswordManagerFormModel)
End Sub

Public Sub UserWantsToOpenUserManagerFrame(ByVal UserManagerFormModel As UserManagerModel)
    'open user manager for the client
    Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameUserManager)
    'reset user manager frame
    Call ResetUserManagerFrame(UserManagerFormModel, OPERATION_NEW)
End Sub

Public Sub UserWantsToOpenPriceFormFrame(ByVal PriceFormFrameModel As PriceFormModel, ByVal operation As FormOperation)
    'open Price Form Interface
    If MainModel.ActiveUserType = USERTYPE_CLIENT Then
        Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.framePriceForm)
    Else
        Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.framePriceForm)
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
    Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameExportUtility)
    'Reset Export Form Frame
    Call ResetExportFormFrame(ExportFormFrameModel)
End Sub

Public Sub UserWantsToOpenDataFormFrame(ByVal DataFormFrameModel As DataFormModel, ByVal ContainerIdentification As DataContainer)
    'open price form interface
    Select Case ContainerIdentification
    
        Case DataContainer.FOR_CLIENTHISTORY
            'open client history interface
            Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.frameRecordsContainer)
            
        Case DataContainer.FOR_PENDINGAPPROVALS
            'open pending list for approver
            Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameRecordsContainer)
            
        Case DataContainer.FOR_ALLHISTORY
            'open client history interface
            Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameRecordsContainer)
            
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
        Call ExtendedMethods.SetStateofControlsToNullState(.txtUsername, .txtPassword)
        'set focus
        .txtUsername.SetFocus
    End With
End Sub

Private Sub ResetPasswordManagerFrame(ByVal PasswordManagerFormModel As PasswordManagerModel)
    With Me
        'Attach Model
        If PasswordModel Is Nothing Then Set PasswordModel = PasswordManagerFormModel
        'clear values of Password manager frame fields
        Call ExtendedMethods.SetStateofControlsToNullState(.txtCurrentPassword, .txtNewPassword, .txtConfirmNewPassword)
        'set focus
        .txtCurrentPassword.SetFocus
    End With
End Sub

Private Sub ResetUserManagerFrame(ByVal UserManagerFormModel As UserManagerModel, ByVal operation As FormOperation)
    With Me
        'Attach Model
        If UserModel Is Nothing Then Set UserModel = UserManagerFormModel
        'clear values of user manager frame fields
        Call ExtendedMethods.SetStateofControlsToNullState(.txtSetUsername, .txtSetPassword, .cmbUserStatus, .cmbUserType, .txtUserEmail, lstUsers)
        'Repopulate ComboBoxes and Listbox
        .cmbUserStatus.List = UserModel.userStatusList
        .cmbUserType.List = UserModel.userTypesList
        With .lstUsers
            .ColumnCount = 7
            .ColumnWidths = "0;45;60"
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
        Call ExtendedMethods.SetStateofControlsToNullState(.lblMainRecordStatus, .txtConditionType, .cmbSalesOrganization, .cmbDistributionChannel, .txtCustomerID, .txtMaterialID, .txtPrice, .cmbCurrency, .txtPriceUnit, .cmbUnitOfMeasure, .txtValidFrom, .txtValidTo)
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
        Call ExtendedMethods.SetStateofControlsToNullState(.lstRecordsContainer, .cmbColumns, .cmbValues)
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
        Else
            DataModel.IsApprover = False
        End If
    'reformat Listbox column with appropriete types
        'Edit Change Date
        Call ExtendedMethods.ReformatListBoxColumns(.lstRecordsContainer, MainTableFields.COL_MAIN_statusChangeDate, TYPE_DATE)
        'price column
        Call ExtendedMethods.ReformatListBoxColumns(.lstRecordsContainer, MainTableFields.COL_MAIN_price, TYPE_CURRENCY)
        'From Date Column
        Call ExtendedMethods.ReformatListBoxColumns(.lstRecordsContainer, MainTableFields.COL_MAIN_validFromDate, TYPE_DATE)
        'To Date Column
        Call ExtendedMethods.ReformatListBoxColumns(.lstRecordsContainer, MainTableFields.COL_MAIN_validToDate, TYPE_DATE)
        'State of Controls of Data Form
        .cmdEditRecord.Enabled = False
    End With
End Sub

Private Sub ResetExportFormFrame(ByVal ExportFormFrameModel As ExportFormModel)
    With Me
        'Attach Model
        If ExportModel Is Nothing Then Set ExportModel = ExportFormFrameModel
        'Clear Data Form Controls
        Call ExtendedMethods.SetStateofControlsToNullState(.txtDateFrom, .txtDateTo, .cmbCustomerID, .cmbUserID, .cmbStatus)
        'repopulate comboboxes
        .cmbCustomerID.List = ExportModel.customerIDsList
        .cmbUserID.List = ExportModel.userIDsList
        .cmbStatus.List = ExportModel.statusesList
        'update model
        Call ExportModel.SetPropertiesToDefaultState
        'input field state
        .txtDateFrom.Value = VBA.Format$(ExportModel.FromDate, DATEFORMAT_FRONTEND)
        .txtDateTo.Value = VBA.Format$(ExportModel.ToDate, DATEFORMAT_FRONTEND)
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
    Call ExtendedMethods.ShowMessage(message, typeOfMessage)
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
            Call ExtendedMethods.ShowMessage("Not authorized to LOGIN! Please contact business to know more details.", TYPE_CRITICAL)
        End If
    Else
        Call ExtendedMethods.ShowMessage(response, TYPE_CRITICAL)
    End If
End Sub

Public Sub ApplicationWantsToUpdateValueListComboBox()
    Me.cmbValues.List = DataModel.ValuesList
    Me.cmbValues.SetFocus
End Sub

Public Sub UserWantsToFilterAndSortDataFormList()
    With Me
        'Clear Data Form Controls
        Call ExtendedMethods.SetStateofControlsToNullState(.lstRecordsContainer)
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
    End With
End Sub

'-------------------------------------------------------------------------
'Button Clicked Operations from Main Frame
'-------------------------------------------------------------------------

'Login Frame

Private Sub OpenNextInterfaceAfterSuccessfulLogin()
    'Open Frame based on client type
    If LoginModel.UserType = USERTYPE_CLIENT Then
        Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.frameWelcome)
    Else
        Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameWelcome)
    End If
    'Update Active User Frame
    With LoginModel
        Call UpdateActiveUserInfomation(.UserName, .UserType, .userStatus, .userID, .Password)
    End With
    'Update Welcome Frame with Username
    Call UpdateWelcomeFrame
End Sub

'Password Manager Frame

Public Sub AfterChangePasswordOperation()
    MsgBox "Password has been changed successfully! Please Sign-In again.", vbInformation, SIGN
    'Go back to logout state
    Call ExtendedMethods.ActivateFrames(Me.frameLogin, Me.frameWelcome)
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
        .txtValidFrom.Value = VBA.Format$(PriceModel.validFromDate, DATEFORMAT_FRONTEND)
        .txtValidTo.Value = VBA.Format$(PriceModel.validToDate, DATEFORMAT_FRONTEND)
        'Hide Buttons
        If MainModel.ActiveUserType = USERTYPE_APPROVER Then
            Call ShowApprovalRejectionButtons(True)
            Call ExtendedMethods.FormEditingState(False, .txtConditionType, .cmbSalesOrganization, .cmbDistributionChannel, .txtCustomerID, .txtMaterialID, .txtPrice, .cmbCurrency, .txtPriceUnit, .cmbUnitOfMeasure, .txtValidFrom, .txtValidTo)
        Else
            Call ShowApprovalRejectionButtons(False)
            Call ExtendedMethods.FormEditingState(True, .txtConditionType, .cmbSalesOrganization, .cmbDistributionChannel, .txtCustomerID, .txtMaterialID, .txtPrice, .cmbCurrency, .txtPriceUnit, .cmbUnitOfMeasure, .txtValidFrom, .txtValidTo)
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
        .txtValidFrom.Value = VBA.Format$(PriceModel.validFromDate, DATEFORMAT_FRONTEND)
        .txtValidTo.Value = VBA.Format$(PriceModel.validToDate, DATEFORMAT_FRONTEND)
        'Hide Buttons & Form Lock Decision
        If MainModel.ActiveUserType = USERTYPE_APPROVER Then
            Call ShowApprovalRejectionButtons(True)
            Call ExtendedMethods.FormEditingState(False, .txtConditionType, .cmbSalesOrganization, .cmbDistributionChannel, .txtCustomerID, .txtMaterialID, .txtPrice, .cmbCurrency, .txtPriceUnit, .cmbUnitOfMeasure, .txtValidFrom, .txtValidTo)
            'Other Buttons State
            .cmdAddNewRecord.Enabled = False
            .cmdUpdateRecord.Enabled = False
            .cmdDeleteRecord.Enabled = False
            .cmdResetPriceForm.Enabled = False
        Else
            Call ShowApprovalRejectionButtons(False)
            Call ExtendedMethods.FormEditingState(True, .txtConditionType, .cmbSalesOrganization, .cmbDistributionChannel, .txtCustomerID, .txtMaterialID, .txtPrice, .cmbCurrency, .txtPriceUnit, .cmbUnitOfMeasure, .txtValidFrom, .txtValidTo)
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
        .cmbUserType.Value = UserModel.UserType
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
        .cmbUserType.Value = UserModel.UserType
        .txtSetUsername.Value = UserModel.UserName
        .txtSetPassword.Value = UserModel.userPassword
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
        With ExtendedMethods
            Call .ChangeControlProperties(Me.lblWelcomeMessage, MESSAGE_WELCOMESCREEN_LOGOUT_STATE)
            Call .SetStateofControlsToNullState(Me.lblActiveUsername, Me.lblActiveUserType, Me.lblActiveUserStatus, Me.lblActiveUserID, Me.lblActiveUserPassword)
        End With
    Else
        'Update Welcome Message While User is Still Logged In
        Call ExtendedMethods.ChangeControlProperties(Me.lblWelcomeMessage, MESSAGE_WELCOMESCREEN_LOGIN_STATE & Me.lblActiveUsername.Caption)
    End If
End Sub

Private Sub UpdateActiveUserInfomation(ByVal uname As String, ByVal uType As String, ByVal uStatus As String, ByVal uID As String, ByVal uPassword As String)
    'Show Active user info on Always On Frame
    With ExtendedMethods
        Call .ChangeControlProperties(Me.lblActiveUsername, uname)
        Call .ChangeControlProperties(Me.lblActiveUserType, uType)
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
        .ActiveUserName = uname
        .ActiveUserPassword = uPassword
        .ActiveUserStatus = uStatus
        .ActiveUserType = uType
    End With
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

Private Sub UserForm_Terminate()
    Set ExtendedMethods = Nothing
    Set MainModel = Nothing
    Set LoginModel = Nothing
    Set UserModel = Nothing
    Set PriceModel = Nothing
    Set DataModel = Nothing
    Set ExportModel = Nothing
End Sub
