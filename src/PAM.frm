VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PAM 
   Caption         =   "Price Approval Manager V1.0"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2535
   OleObjectBlob   =   "PAM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'-------------------------------------------------------------------------
'Public Events
'-------------------------------------------------------------------------

'Main Frame Events
Public Event OpenLoginFrame()
Public Event ExitApp()
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

'-------------------------------------------------------------------------
'SETTINGS
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
End Type

Private this As TViewComponents

'-------------------------------------------------------------------------
'Private Variables/Objects
'-------------------------------------------------------------------------

Private ExtendedMethods As MultiFrameViewExtended

'-------------------------------------------------------------------------
'Properties
'-------------------------------------------------------------------------

Private Property Get MainModel() As AppModel
    Set MainModel = this.MainModel
End Property

Private Property Set MainModel(ByVal vNewValue As AppModel)
    Set this.MainModel = vNewValue
End Property

Private Property Get LoginModel() As LoginFormModel
    Set LoginModel = this.LoginModel
End Property

Private Property Set LoginModel(ByVal vNewValue As LoginFormModel)
    Set this.LoginModel = vNewValue
End Property

Private Property Get PasswordModel() As PasswordManagerModel
    Set PasswordModel = this.PasswordModel
End Property

Private Property Set PasswordModel(ByVal vNewValue As PasswordManagerModel)
    Set this.PasswordModel = vNewValue
End Property

Private Property Get UserModel() As UserManagerModel
    Set UserModel = this.UserModel
End Property

Private Property Set UserModel(ByVal vNewValue As UserManagerModel)
    Set this.UserModel = vNewValue
End Property

Private Property Get PriceModel() As PriceFormModel
    Set PriceModel = this.PriceModel
End Property

Private Property Set PriceModel(ByVal vNewValue As PriceFormModel)
    Set this.PriceModel = vNewValue
End Property

'-------------------------------------------------------------------------
'public Methods Called From Presenters
'-------------------------------------------------------------------------

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

Public Sub UserWantsToOpenPriceFormFrame(ByVal PriceFormFrameModel As PriceFormModel)
    'open Price Form Interface
    Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.framePriceForm)
    'Reset Price Form Frame
    'Call ResetPriceFormFrame(PriceFormFrameModel, OPERATION_NEW)
End Sub

'------------------------------------------------------------------------
'Public methods to perfrom operations
'-------------------------------------------------------------------------

Public Sub UserWantsToUpdateUserManagerRecord()
    'reset user manager frame
    Call ResetUserManagerFrame(UserModel, OPERATION_UPDATE)
End Sub

Public Sub ShowWarning(ByVal message As String, ByVal typeOfMessage As messageType)
    Call ExtendedMethods.ShowMessage(message, TYPE_CRITICAL)
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

'-------------------------------------------------------------------------
'User Form Events
'-------------------------------------------------------------------------

Private Sub cmdCancelFromLoginInterface_Click()
    RaiseEvent CloseLoginFrame
End Sub

Private Sub cmdCancelPasswordManager_Click()
    RaiseEvent ClosePasswordManagerFrame
End Sub

Private Sub cmdOpenPriceForm_Click()
    RaiseEvent OpenPriceFormFrame
End Sub

Private Sub cmdOpenUserManager_Click()
    RaiseEvent OpenUserManagerFrame
End Sub

Private Sub cmdResetUserManager_Click()
    RaiseEvent ResetUserManagerFrame
End Sub

Private Sub cmdUpdatePassword_Click()
    RaiseEvent ChangePassword
End Sub

Private Sub cmdExit_Click()
    RaiseEvent ExitApp
End Sub

Private Sub cmdLogin_Click()
    RaiseEvent Login
End Sub

Private Sub cmdOpenLoginInterface_Click()
    RaiseEvent OpenLoginFrame
End Sub

Private Sub cmdOpenPasswordManager_Click()
    RaiseEvent OpenPasswordManagerFrame
End Sub

Private Sub cmdCancelUserManager_Click()
    RaiseEvent CloseUserManagerFrame
End Sub

Private Sub cmdAddNewUser_Click()
    RaiseEvent DoCRUDOperationForUserManager(CRUD_OPERATION_ADDNEW)
End Sub

Private Sub cmdDeleteUser_Click()
    RaiseEvent DoCRUDOperationForUserManager(CRUD_OPERATION_DELETE)
End Sub

Private Sub cmdUpdateUser_Click()
    RaiseEvent DoCRUDOperationForUserManager(CRUD_OPERATION_UPDATE)
End Sub

Private Sub cmdApproverLogout_Click()
    'Logout State
    Call ExtendedMethods.ActivateFrames(Me.frameLogin, Me.frameWelcome)
    Call UpdateWelcomeFrame(FORM_LOGIN)
End Sub

Private Sub cmdCancelExportUtility_Click()
    'cancel
    Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameWelcome)
    Call UpdateWelcomeFrame
End Sub

Private Sub cmdCancelPriceFormInterface_Click()
    'back to the dashboard
    Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.frameWelcome)
    Call UpdateWelcomeFrame
End Sub

Private Sub cmdCancelRecordContainer_Click()
    'back to the dashboard
    With ExtendedMethods
        If Me.lblActiveUserType.Caption = USERTYPE_CLIENT Then
            Call .ActivateFrames(Me.frameClient, Me.frameWelcome)
        Else
            Call .ActivateFrames(Me.frameApprover, Me.frameWelcome)
        End If
    End With
    Call UpdateWelcomeFrame
End Sub

Private Sub cmdClientLogout_Click()
    'Logout State
    Call ExtendedMethods.ActivateFrames(Me.frameLogin, Me.frameWelcome)
    Call UpdateWelcomeFrame(FORM_LOGIN)
End Sub

Private Sub cmdEditRecord_Click()
    'show
    'back to the dashboard
    With ExtendedMethods
        If Me.lblActiveUserType.Caption = USERTYPE_CLIENT Then
            Call .ActivateFrames(Me.frameClient, Me.frameWelcome)
        Else
            Call .ActivateFrames(Me.frameApprover, Me.frameWelcome)
        End If
    End With
End Sub

Private Sub cmdOpenAllHistory_Click()
    'open client history interface
    Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameRecordsContainer)
End Sub

Private Sub cmdOpenClientHistory_Click()
    'open client history interface
    Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.frameRecordsContainer)
End Sub

Private Sub cmdOpenExportUtility_Click()
    'open export utility
    Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameExportUtility)
End Sub

Private Sub cmdOpenPendingList_Click()
    'open pending list for approver
    Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameRecordsContainer)
End Sub

'-------------------------------------------------------------------------
'User Manager Fileds Change Events
'-------------------------------------------------------------------------

Private Sub cmbUserStatus_Change()
    'Hydrate model property
    UserModel.userStatus = Me.cmbUserStatus.value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.cmbUserStatus, UserModel.IsValidField(COL_userStatus), TYPE_NA
End Sub

Private Sub cmbUserType_Change()
    'Hydrate model property
    UserModel.userType = Me.cmbUserType.value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.cmbUserType, UserModel.IsValidField(COL_userType), TYPE_NA
End Sub

Private Sub txtSetUsername_Change()
    'Hydrate model property
    UserModel.userName = Me.txtSetUsername.value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtSetUsername, UserModel.IsValidField(COL_userName), TYPE_FIXEDLENGTHSTRING, 6
End Sub

Private Sub txtSetPassword_Change()
    'hydrate model property
    UserModel.userPassword = Me.txtSetPassword.value
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtSetPassword, UserModel.IsValidField(COL_password), TYPE_WRONGPASSWORDPATTERN
End Sub

Private Sub lstUsers_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With Me.lstUsers
        If .ListIndex > 0 Then
            If .ListIndex > 2 Then
                'hydrate model property
                UserModel.userIndex = .List(.ListIndex, 0) + 1
                'Update Record
                RaiseEvent UpdateUserManagerFrameRecord
            Else
                'Just for the safetly that they couldn't be able to edit dev's information
                Call ExtendedMethods.ShowMessage("You are not allowed to Update them!", TYPE_INFORMATION)
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
    Me.txtConfirmNewPassword.value = vbNullString
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
    LoginModel.password = Me.txtPassword.Text
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtPassword, LoginModel.IsValidPassword, TYPE_NA
End Sub

Private Sub txtUsername_Change()
    'hydrate model property
    LoginModel.userName = Me.txtUsername.Text
    'Validate Field
    ExtendedMethods.UpdateControlAfterValidation Me.txtUsername, LoginModel.IsValidUsername, TYPE_NA
End Sub

'-------------------------------------------------------------------------
'Private methods
'-------------------------------------------------------------------------

Public Sub InItApplication(ByVal ApplicationModel As AppModel)
    'init Extended Methods
    Set MainModel = ApplicationModel
    Set ExtendedMethods = New MultiFrameViewExtended
    With ExtendedMethods
        'Re-Dimension UserForm
        Set .TargetForm = Me
        .formWidth = 600
        .formHeight = 360
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
        .sideFrameHeight = 234
        'Main Panel Frames Properties
        .mainFrameTop = 6
        .mainFrameLeft = 152
        .mainFrameWidth = 430
        .mainFrameHeight = 318
        'plug static data sources to the relative comboboxes
        Call .HydrateComboBox(Me.cmbCurrency, modDataSources.arrListofCurrencies)
        Call .HydrateComboBox(Me.cmbUnitOfMeasure, modDataSources.arrListOfUnitOfMeasure)
        'InIt Interface
        Call .ActivateFrames(Me.frameLogin, Me.frameWelcome)
        Call UpdateWelcomeFrame(FORM_LOGIN)
    End With
End Sub

'-------------------------------------------------------------------------
'Frame Reset Methods
'-------------------------------------------------------------------------

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

Private Sub ResetUserManagerFrame(ByVal UserManagerFormModel As UserManagerModel, ByVal Operation As FormOperation)
    With Me
        'Attach Model
        If UserModel Is Nothing Then Set UserModel = UserManagerFormModel
        'clear values of user manager frame fields
        Call ExtendedMethods.SetStateofControlsToNullState(.lblUserID, .txtSetUsername, .txtSetPassword, .cmbUserStatus, .cmbUserType, lstUsers)
        'Repopulate ComboBoxes and Listbox
        .cmbUserStatus.List = UserModel.userStatusList
        .cmbUserType.List = UserModel.userTypesList
        With .lstUsers
            .ColumnCount = 6
            .ColumnWidths = "35;45;60"
            .List = UserModel.usersTable
        End With
        'Put Default Values based on Operation
        If Operation = OPERATION_NEW Then
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

'-------------------------------------------------------------------------
'Button Clicked Operations
'-------------------------------------------------------------------------

'Login Frame

Private Sub OpenNextInterfaceAfterSuccessfulLogin()
    'Open Frame based on client type
    If LoginModel.userType = USERTYPE_CLIENT Then
        Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.frameWelcome)
    Else
        Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameWelcome)
    End If
    'Update Active User Frame
    With LoginModel
        Call UpdateActiveUserInfomation(.userName, .userType, .userStatus, .userID, .password)
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

Public Sub AfterUserManagerCRUDOperation(ByVal TypeOfOperation As CRUDOperations)
    Select Case TypeOfOperation
        Case CRUDOperations.CRUD_OPERATION_ADDNEW
            MsgBox "New USER added successfully!", vbInformation, SIGN
        Case CRUDOperations.CRUD_OPERATION_UPDATE
            MsgBox "User's record has been UPDATED successfully!", vbInformation, SIGN
        Case CRUDOperations.CRUD_OPERATION_DELETE
            MsgBox "User has been DELETED successfully!", vbInformation, SIGN
    End Select
    'Refresh Data Again
    RaiseEvent OpenUserManagerFrame
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
End Sub




























'-------------------------------------------------------------------------
'Abstract Methods
'-------------------------------------------------------------------------

Private Sub StateForNewRecordForUserManager()
    With Me
        'Update Model
        Call UserModel.SetPropertiesToNewUserState
        'Input Field State
        .lblUserID.Caption = UserModel.userID
        .cmbUserStatus.value = UserModel.userStatus
        .cmbUserType.value = UserModel.userType
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
        .lblUserID.Caption = UserModel.userID
        .cmbUserStatus.value = UserModel.userStatus
        .cmbUserType.value = UserModel.userType
        .txtSetUsername.value = UserModel.userName
        .txtSetPassword.value = UserModel.userPassword
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

Private Sub UpdateActiveUserInfomation(ByVal uName As String, ByVal uType As String, ByVal uStatus As String, ByVal uID As String, ByVal uPassword As String)
    'Show Active user info on Always On Frame
    With ExtendedMethods
        Call .ChangeControlProperties(Me.lblActiveUsername, uName)
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
        .ActiveUserName = uName
        .ActiveUserPassword = uPassword
        .ActiveUserStatus = uStatus
        .ActiveUserType = uType
    End With
End Sub
