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

Private WithEvents view As PAM
Attribute view.VB_VarHelpID = -1
Private WithEvents presenterLoginFrame As LoginFormPresenter
Attribute presenterLoginFrame.VB_VarHelpID = -1
Private WithEvents presenterPasswordManagerFrame As PasswordManagerPresenter
Attribute presenterPasswordManagerFrame.VB_VarHelpID = -1
Private WithEvents presenterUserManagerFrame As UserManagerPresenter
Attribute presenterUserManagerFrame.VB_VarHelpID = -1
Private WithEvents presenterPriceFormFrame As PriceFormPresenter
Attribute presenterPriceFormFrame.VB_VarHelpID = -1

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
Public Sub InItApplicationModel(ByVal ListofCurrencies As Variant, _
                                ByVal ListOfUnits As Variant, _
                                ByVal TypesOfUser As Variant, _
                                ByVal StatusOfUser As Variant)
    Set model = New AppModel
    Call model.InItModel(mainService.GetListOfAllRecordsFromContextTable, _
                        userService.GetListOfAllRecordsFromContextTable, _
                        ListofCurrencies, _
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

Private Sub view_OpenPriceFormFrame()
    Call OpenPriceForm
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
    'Hydrate model property
    Call LModel.InItModel(model.usersTable)
    'Open Login Form
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
    'Hydrate Model property
    Call PMModel.InItModel(model.ActiveUserPassword)
    'Open Passwordmanager Frame
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
    'Hydrate Model Properties
    Call UMModel.InItModel(model.usersTable, model.userStatusList, model.userTypeList)
    'Open User manager frame
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
'Columns of Users Table
    '0. index
    '1. userID
    '2. userStatus
    '3. userType
    '4. userName
    '5. password
'index will be alloted automatically so we should start from 1 to 5

Private Sub presenterUserManagerFrame_AddNewUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal UMModel As UserManagerModel)
    Call RefreshUsersTable
    With UMModel
        Call userService.AddNewRecordToContextTable(model.GetNewID(TablesOfThisApplication.TABLE_USERS), .userStatus, .userType, .userName, .userPassword)
    End With
    Call view.AfterUserManagerCRUDOperation(TypeOfOperation, True)
End Sub

Private Sub presenterUserManagerFrame_UpdateUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal UMModel As UserManagerModel)
    Call RefreshUsersTable
    Dim response As Variant
    With UMModel
    response = model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_USERS, .userIndex - 1, .UserID)
        If response = True Then
            Call userService.UpdateRecordOfContextTable(.userIndex - 1, .UserID, .userStatus, .userType, .userName, .userPassword)
            Call view.AfterUserManagerCRUDOperation(TypeOfOperation, True)
        Else
            Call view.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub presenterUserManagerFrame_DeleteUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal UMModel As UserManagerModel)
    Call RefreshUsersTable
    Dim response As Variant
    With UMModel
        response = model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_USERS, .userIndex - 1, .UserID)
        If response = True Then
            Call userService.DeleteRecordofContextTable(.userIndex - 1)
            Call view.AfterUserManagerCRUDOperation(TypeOfOperation, True)
        Else
            Call view.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub presenterUserManagerFrame_OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
    Call view.ShowWarning(message, TYPE_CRITICAL)
End Sub

'-------------------
'PRICE FORM Events
'-------------------

Private Sub presenterPriceFormFrame_OpenPriceForm(ByVal PFModel As PriceFormModel)
    'Hydrate Model property
    Call PFModel.InItModel(model.mainTable, model.CurrenciesList, model.unitOfMeasuresList)
    'Open PriceForm Frame
    Call view.UserWantsToOpenPriceFormFrame(PFModel)
End Sub

Private Sub presenterPriceFormFrame_ResetPriceForm()
    Call OpenPriceForm
End Sub

'Values shouold be in sequence of how they are in table!
'Columns of main table
    '0. index
    '1. recordID
    '2. recordStatus
    '3. statusChangeDate
    '4. customerID
    '5. materialID
    '6. price
    '7. currency
    '8. unitOfPrice
    '9. unitOfMeasure
    '10. validFromDate
    '11. validToDate
'index will be alloted automatically so we should start from 1 to 11

Private Sub presenterPriceFormFrame_AddNewRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
    Call RefreshMainTable
    With PFModel
        Call mainService.AddNewRecordToContextTable(model.GetNewID(TablesOfThisApplication.TABLE_MAINRECORDS), .UserID, .recordStatus, vbNullString, .customerID, .materialID, .price, .currencyType, .unitOfPrice, .unitOfMeasure, .validFromDate, .validToDate)
    End With
    Call view.AfterPriceFormCRUDOperation(TypeOfOperation, True)
End Sub

Private Sub presenterPriceFormFrame_UpdateRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
    Call RefreshMainTable
    Dim response As Variant
    With PFModel
    response = model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_MAINRECORDS, .index - 1, .RecordID)
        If response = True Then
            Call mainService.UpdateRecordOfContextTable(.index - 1, .RecordID, .UserID, .recordStatus, .statusChangeDate, .customerID, .materialID, .price, .currencyType, .unitOfPrice, .unitOfMeasure, .validFromDate, .validToDate)
            Call view.AfterPriceFormCRUDOperation(TypeOfOperation, True)
        Else
            Call view.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub presenterPriceFormFrame_DeleteRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
    Call RefreshMainTable
    Dim response As Variant
    With PFModel
    response = model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_MAINRECORDS, .index - 1, .RecordID)
        If response = True Then
            Call mainService.DeleteRecordofContextTable(.index - 1)
            Call view.AfterPriceFormCRUDOperation(TypeOfOperation, True)
        Else
            Call view.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub presenterPriceFormFrame_ClosePriceForm()
    Call view.UserWantsToCloseFrame(FORM_PRICEFORM)
End Sub

Private Sub presenterPriceFormFrame_OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
    Call view.ShowWarning(message, TYPE_CRITICAL)
End Sub

'----------------------------------------------------------------------------------------
'Following two event handllers are for Approving and Rejecting Processes!!
'technically, Approving and Rejecting Processes are exactly UPDATE process!
'we can also combine them in one procedure but I don't want because
'if in future, business demands something more with Approval and Rejection processes then
'they can get integrate model's data from here and process further!
'E.g. Like sending emails, sending reports, notify client through SMS etc etc.
'any thing can integrate from this point
'----------------------------------------------------------------------------------------

Private Sub presenterPriceFormFrame_ApproveRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
    Call RefreshMainTable
    Dim response As Variant
    With PFModel
    response = model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_MAINRECORDS, .index - 1, .RecordID)
        If response = True Then
            Call mainService.UpdateRecordOfContextTable(.index - 1, .RecordID, .UserID, .recordStatus, .statusChangeDate, .customerID, .materialID, .price, .currencyType, .unitOfPrice, .unitOfMeasure, .validFromDate, .validToDate)
            Call view.AfterPriceFormCRUDOperation(TypeOfOperation, True)
        Else
            Call view.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub presenterPriceFormFrame_RejectRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
    Call RefreshMainTable
    Dim response As Variant
    With PFModel
    response = model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_MAINRECORDS, .index - 1, .RecordID)
        If response = True Then
            Call mainService.UpdateRecordOfContextTable(.index - 1, .RecordID, .UserID, .recordStatus, .statusChangeDate, .customerID, .materialID, .price, .currencyType, .unitOfPrice, .unitOfMeasure, .validFromDate, .validToDate)
            Call view.AfterPriceFormCRUDOperation(TypeOfOperation, True)
        Else
            Call view.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

'-------------------------------------------------------------------------
'Private Methods to increase abstractiveness of presenter code
'-------------------------------------------------------------------------

Private Sub OpenLoginForm()
    'Refresh Model Data Table with Updated information
    Call RefreshUsersTable '(Only user because Login Mechanism does not required other tables)
    'Open New Login Form
    If presenterLoginFrame Is Nothing Then Set presenterLoginFrame = New LoginFormPresenter
    'InIt Presenter from here
    Call presenterLoginFrame.InItLoginForm(view)
End Sub

Private Sub OpenPasswordManager()
    'Refresh Model Data Table with Updated information
    Call RefreshUsersTable '(Only user because password manager does not required other tables)
    'Open Password Manager
    If presenterPasswordManagerFrame Is Nothing Then Set presenterPasswordManagerFrame = New PasswordManagerPresenter
    'InIt Presenter from here
    Call presenterPasswordManagerFrame.InItPasswordManagerForm(view)
End Sub

Private Sub OpenUserManager()
    'Refresh Model Data Table with Updated information
    Call RefreshUsersTable '(Only user because password manager does not required other tables)
    'open usermanager
    If presenterUserManagerFrame Is Nothing Then Set presenterUserManagerFrame = New UserManagerPresenter
    'InIt Presenter from here
    With model
        Call presenterUserManagerFrame.InItUserManagerForm(view)
    End With
End Sub

Private Sub OpenPriceForm()
    'Refresh Model Data Table with Updated information
    Call RefreshMainTable '(Only user because password manager does not required other tables)
    'open usermanager
    If presenterPriceFormFrame Is Nothing Then Set presenterPriceFormFrame = New PriceFormPresenter
    'InIt Presenter from here
    With model
        Call presenterPriceFormFrame.InItPriceForm(view)
    End With
End Sub

Private Sub RefreshMainTable()
    model.mainTable = mainService.GetListOfAllRecordsFromContextTable
End Sub

Private Sub RefreshUsersTable()
    model.usersTable = userService.GetListOfAllRecordsFromContextTable
End Sub

'-------------------------------------------------------------------------
'Class Events
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set presenterPriceFormFrame = Nothing
    Set presenterUserManagerFrame = Nothing
    Set presenterPasswordManagerFrame = Nothing
    Set presenterLoginFrame = Nothing
    Set view = Nothing
    Set model = Nothing
    Set userService = Nothing
    Set mainService = Nothing
End Sub



