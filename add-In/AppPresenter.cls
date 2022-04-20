VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------
'Event handllers
'-------------------------------------------------------------------------

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents view As PAM
Attribute view.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents presenterLoginFrame As LoginFormPresenter
Attribute presenterLoginFrame.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents presenterPasswordManagerFrame As PasswordManagerPresenter
Attribute presenterPasswordManagerFrame.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents presenterUserManagerFrame As UserManagerPresenter
Attribute presenterUserManagerFrame.VB_VarHelpID = -1

'-------------------------------------------------------------------------
'Private Componenets Set
'-------------------------------------------------------------------------

Private Type TAppPresenterComponents
    mainService As IServices
    userService As IServices
    model As AppModel
End Type

Private this As TAppPresenterComponents

'-------------------------------------------------------------------------
'Private properties
'-------------------------------------------------------------------------

Private Property Get model() As AppModel
    Set model = this.model
End Property

Private Property Set model(ByVal vNewValue As AppModel)
    Set this.model = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get mainService() As IServices
    Set mainService = this.mainService
End Property

Private Property Set mainService(ByVal vNewValue As IServices)
    Set this.mainService = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get userService() As IServices
    Set userService = this.userService
End Property

Private Property Set userService(ByVal vNewValue As IServices)
    Set this.userService = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt Application/Starting Point of Application at pattern level
'-------------------------------------------------------------------------

'InIt Main Services
Public Sub InItMainService(ByVal SelectedRepositoryType As RepositoryType, ByVal TableName As String)
    
    Set mainService = New ServicesMain
    Call mainService.InItService(SelectedRepositoryType, TableName)
End Sub

'InIt User Services
Public Sub InItUserService(ByVal SelectedRepositoryType As RepositoryType, ByVal TableName As String)
    
    Set userService = New ServicesUser
    Call userService.InItService(SelectedRepositoryType, TableName)
End Sub

'InIt Model
Public Sub InItApplicationModel(ByVal ListOfCurrencies As Variant, _
                                ByVal ListOfUnits As Variant, _
                                ByVal TypesOfUser As Variant, _
                                ByVal StatusOfUser As Variant)
    Set model = New AppModel
    Call model.InItModel(mainService.GetListOfAllRecordsFromContextTable, _
                        userService.GetListOfAllRecordsFromContextTable, _
                        ListOfCurrencies, _
                        ListOfUnits, _
                        TypesOfUser, _
                        StatusOfUser)
End Sub

'InIt Main View / This step will open Main Interface
Public Sub InItApp()
    Set view = New PAM
    With view
        Call .InItApplication(model) 'For Supervision View
        Call .Show
    End With
End Sub

'-------------------------------------------------------------------------
'Frame Opening Operations / Side Panel Buttons
'-------------------------------------------------------------------------

Private Sub view_OpenLoginFrame()
    Call OpenLoginForm
End Sub

Private Sub view_OpenPasswordManagerFrame()
    Call OpenPasswordManager
End Sub

Private Sub view_OpenUserManagerFrame()
    Call OpenUserManager
End Sub

Private Sub view_ExitApp()
    Call view.OnCancel
End Sub

'-------------------------------------------------------------------------
'Frame Operations
'-------------------------------------------------------------------------


'------------------
'Login Frame Events
'------------------

Private Sub presenterLoginFrame_OpenLoginForm(ByVal LModel As LoginFormModel)
    Call view.UserWantsToOpenLoginFrame(LModel)
End Sub

Private Sub presenterLoginFrame_Login()
    Call view.UserWantsToLogin
End Sub

Private Sub presenterLoginFrame_CancelLoginForm()
    Call view.UserWantsToCloseFrame(FORM_LOGIN)
End Sub

'------------------------
'Password Manager Events
'------------------------

Private Sub presenterPasswordManagerFrame_OpenPasswordManagerForm(ByVal PMModel As PasswordManagerModel)
    Call view.UserWantsToOpenPasswordManagerFrame(PMModel)
End Sub

Private Sub presenterPasswordManagerFrame_ChangePassword(ByVal newPassword As String)
    'index, userID, userStatus, userType, userName, password
    With model
        Call userService.UpdateRecordOfContextTable(.ActiveUserIndex, .ActiveUserID, .ActiveUserStatus, .ActiveUserType, .ActiveUserName, newPassword)
        Call view.AfterChangePasswordOperation
    End With
End Sub

Private Sub presenterPasswordManagerFrame_OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
    Call view.ShowWarning(message, typeOfMessage)
End Sub

Private Sub presenterPasswordManagerFrame_CancelPasswordManagerForm()
    Call view.UserWantsToCloseFrame(FORM_PASSWORDMANAGER)
End Sub

'-------------------
'User Manager Events
'-------------------

Private Sub presenterUserManagerFrame_OpenUserManagerForm(ByVal UMModel As UserManagerModel)
    Call view.UserWantsToOpenUserManagerFrame(UMModel)
End Sub

Private Sub presenterUserManagerFrame_CloseUserManagerForm()
    Call view.UserWantsToCloseFrame(FORM_USERMANAGER)
End Sub

Private Sub presenterUserManagerFrame_ResetUserManagerForm()
    Call OpenUserManager
End Sub

Private Sub presenterUserManagerFrame_UpdateUserManagerFormRecord()
    Call view.UserWantsToUpdateUserManagerRecord
End Sub

'Values shouold be in sequence of how they are in table!
'usersTable = 1. index   2. userID  3. userStatus  4. userType    5. userName    6. password
'index will be alloted automatically so we should start from 2 to 6

Private Sub presenterUserManagerFrame_AddNewUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal userID As String, ByVal userStatus As String, ByVal userType As String, ByVal userName As String, ByVal userPassword As String)
    Call userService.AddNewRecordToContextTable(userID, userStatus, userType, userName, userPassword)
    Call view.AfterUserManagerCRUDOperation(TypeOfOperation)
End Sub

Private Sub presenterUserManagerFrame_UpdateUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal index As Long, ByVal userID As String, ByVal userStatus As String, ByVal userType As String, ByVal userName As String, ByVal userPassword As String)
    Call userService.UpdateRecordOfContextTable(index, userID, userStatus, userType, userName, userPassword)
    Call view.AfterUserManagerCRUDOperation(TypeOfOperation)
End Sub

Private Sub presenterUserManagerFrame_DeleteUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal index As Long)
    Call userService.DeleteRecordofContextTable(index)
    Call view.AfterUserManagerCRUDOperation(TypeOfOperation)
End Sub

Private Sub presenterUserManagerFrame_OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
    Call view.ShowWarning(message, TYPE_CRITICAL)
End Sub

'-------------------------------------------------------------------------
'Private Methods to increase abstractiveness of presenter code
'-------------------------------------------------------------------------

Private Sub OpenLoginForm()
    'Refresh Model Data Table with Updated information
    Call RefreshUsersTable '(Only user because Login Mechanism does not required other tables)
    model.usersTable = userService.GetListOfAllRecordsFromContextTable
    'Open New Login Form
    If presenterLoginFrame Is Nothing Then Set presenterLoginFrame = New LoginFormPresenter
    'InIt Presenter from here
    Call presenterLoginFrame.InItLoginForm(view, model.usersTable)
End Sub

Private Sub OpenPasswordManager()
    'Refresh Model Data Table with Updated information
    Call RefreshUsersTable '(Only user because password manager does not required other tables)
    'Open Password Manager
    If presenterPasswordManagerFrame Is Nothing Then Set presenterPasswordManagerFrame = New PasswordManagerPresenter
    'InIt Presenter from here
    Call presenterPasswordManagerFrame.InItPasswordManagerForm(view, model.ActiveUserPassword)
End Sub

Private Sub OpenUserManager()
    'Refresh Model Data Table with Updated information
    Call RefreshUsersTable '(Only user because password manager does not required other tables)
    'open usermanager
    If presenterUserManagerFrame Is Nothing Then Set presenterUserManagerFrame = New UserManagerPresenter
    'InIt Presenter from here
    With model
        Call presenterUserManagerFrame.InItUserManagerForm(view, .usersTable, .userStatusList, .userTypeList)
    End With
End Sub

Private Sub RefreshMainTable()
    model.MainTable = mainService.GetListOfAllRecordsFromContextTable
End Sub

Private Sub RefreshUsersTable()
    model.usersTable = userService.GetListOfAllRecordsFromContextTable
End Sub

'-------------------------------------------------------------------------
'Class Events
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set presenterUserManagerFrame = Nothing
    Set presenterPasswordManagerFrame = Nothing
    Set presenterLoginFrame = Nothing
    Set view = Nothing
    Set model = Nothing
    Set userService = Nothing
    Set mainService = Nothing
End Sub


