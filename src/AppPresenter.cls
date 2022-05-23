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
Private WithEvents presenterDataFormFrame As DataFormPresenter
Attribute presenterDataFormFrame.VB_VarHelpID = -1
Private WithEvents presenterExportFormFrame As ExportFormPresenter
Attribute presenterExportFormFrame.VB_VarHelpID = -1

'-------------------------------------------------------------------------
'Private Componenets Set
'-------------------------------------------------------------------------

Private Type TAppPresenterComponents
    databaseConnectionStatus As Boolean
    mainService As IServices
    userService As IServices
    model As AppModel
    report As DocumentGeneratorExcel
End Type

Private this As TAppPresenterComponents

'-------------------------------------------------------------------------
'Private properties
'-------------------------------------------------------------------------

Public Property Get databaseConnectionStatus() As Boolean
    databaseConnectionStatus = this.databaseConnectionStatus
End Property

Public Property Let databaseConnectionStatus(ByVal RHS As Boolean)
    this.databaseConnectionStatus = RHS
End Property

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

Private Property Get report() As DocumentGeneratorExcel
    Set report = this.report
End Property

Private Property Set report(ByVal vNewValue As DocumentGeneratorExcel)
    Set this.report = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt Application/Starting Point of Application at pattern level
'-------------------------------------------------------------------------

'InIt Main Services
Public Sub InItMainService(ByVal SelectedRepositoryType As RepositoryType, _
                            ByVal TableName As String, _
                            ByVal TableHeaders As Variant, _
                            ByVal ConnectionString As String)
    Set mainService = New ServicesMain
    Call mainService.InItService(SelectedRepositoryType, TableName, TableHeaders, ConnectionString)
    databaseConnectionStatus = mainService.IsEverythingOkayInEngine
End Sub

'InIt User Services
Public Sub InItUserService(ByVal SelectedRepositoryType As RepositoryType, _
                            ByVal TableName As String, _
                            ByVal TableHeaders As Variant, _
                            ByVal ConnectionString As String)
    Set userService = New ServicesUser
    Call userService.InItService(SelectedRepositoryType, TableName, TableHeaders, ConnectionString)
    databaseConnectionStatus = userService.IsEverythingOkayInEngine
End Sub

'InIt Model
Public Sub InItApplicationModel(ByVal ListofCurrencies As Variant, _
                                ByVal ListOfUnits As Variant, _
                                ByVal TypesOfUser As Variant, _
                                ByVal StatusOfUser As Variant, _
                                ByVal StatusOfRecord As Variant, _
                                ByVal ListOfSalesOrganizations As Variant, _
                                ByVal ListofDistrobutionChannels As Variant)
    Set model = New AppModel
    Call model.InItModel(mainService.GetListOfAllRecordsFromContextTable, _
                        userService.GetListOfAllRecordsFromContextTable, _
                        ListofCurrencies, _
                        ListOfUnits, _
                        TypesOfUser, _
                        StatusOfUser, _
                        StatusOfRecord, _
                        ListOfSalesOrganizations, _
                        ListofDistrobutionChannels)
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
    Call OpenPriceForm(OPERATION_NEW)
End Sub

Private Sub view_OpenDataFormFrame(ByVal ContainerIdentifier As DataContainer)
    Select Case ContainerIdentifier
        Case DataContainer.FOR_CLIENTHISTORY
            Call OpenDataForm(FOR_CLIENTHISTORY)
        Case DataContainer.FOR_PENDINGAPPROVALS
            Call OpenDataForm(FOR_PENDINGAPPROVALS)
        Case DataContainer.FOR_ALLHISTORY
            Call OpenDataForm(FOR_ALLHISTORY)
    End Select
End Sub

Private Sub view_OpenExportFormFrame()
    Call OpenExportForm
End Sub

Private Sub view_Logout()
    Call view.UserWantsToLogout
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
        If userService.IsEverythingOkayInEngine Then
            Call view.AfterChangePasswordOperation
        End If
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
        Call userService.AddNewRecordToContextTable(model.GetNewID(TablesOfThisApplication.TABLE_USERS), .userStatus, .UserType, .UserName, .userPassword)
    End With
    If userService.IsEverythingOkayInEngine Then
        Call view.AfterUserManagerCRUDOperation(TypeOfOperation, True)
    End If
End Sub

Private Sub presenterUserManagerFrame_UpdateUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal UMModel As UserManagerModel)
    Call RefreshUsersTable
    Dim response As Variant
    With UMModel
    response = model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_USERS, .userIndex, .userID)
        If response = True Then
            Call userService.UpdateRecordOfContextTable(.userIndex, .userID, .userStatus, .UserType, .UserName, .userPassword)
            If userService.IsEverythingOkayInEngine Then
                Call view.AfterUserManagerCRUDOperation(TypeOfOperation, True)
            End If
        Else
            Call view.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub presenterUserManagerFrame_DeleteUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal UMModel As UserManagerModel)
    Call RefreshUsersTable
    Dim response As Variant
    With UMModel
        response = model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_USERS, .userIndex, .userID)
        If response = True Then
            Call userService.DeleteRecordofContextTable(.userIndex)
            If userService.IsEverythingOkayInEngine Then
                Call view.AfterUserManagerCRUDOperation(TypeOfOperation, True)
            End If
        Else
            Call view.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub presenterUserManagerFrame_OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
    Call view.ShowWarning(message, typeOfMessage)
End Sub

'-------------------
'PRICE FORM Events
'-------------------

Private Sub presenterPriceFormFrame_OpenPriceForm(ByVal PFModel As PriceFormModel, ByVal operation As FormOperation)
    If operation = OPERATION_NEW Then
        'Hydrate Model property
        With model
            Call PFModel.InItModel(.mainTable, .CurrenciesList, .unitOfMeasuresList, .salesOrganizationList, .distributionChannelsList)
        End With
        'Open PriceForm Frame
        Call view.UserWantsToOpenPriceFormFrame(PFModel, OPERATION_NEW)
    Else
        With model
            Call PFModel.InItModel(.mainTable, .CurrenciesList, .unitOfMeasuresList, .salesOrganizationList, .distributionChannelsList, .CurrentUpdateIndex)
        End With
        'Open PriceForm Frame
        Call view.UserWantsToOpenPriceFormFrame(PFModel, OPERATION_UPDATE)
    End If
End Sub

Private Sub presenterPriceFormFrame_ResetPriceForm()
    Call OpenPriceForm(OPERATION_NEW)
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
    'I added here StatusChangeDate, because, While filtering the date it is causing an issue! Consider this statusChangeDate field as record added/update date!
    With PFModel
        Call mainService.AddNewRecordToContextTable(model.GetNewID(TablesOfThisApplication.TABLE_MAINRECORDS), _
                                                    .userID, _
                                                    .recordStatus, _
                                                    VBA.Format$(.statusChangeDate, DATEFORMAT_BACKEND), _
                                                    .conditionType, _
                                                    .salesOrganization, _
                                                    .distributionChannel, _
                                                    .customerID, _
                                                    .materialID, _
                                                    .price, _
                                                    .currencyType, _
                                                    .unitOfPrice, _
                                                    .unitOfMeasure, _
                                                    VBA.Format$(.validFromDate, DATEFORMAT_BACKEND), _
                                                    VBA.Format$(.validToDate, DATEFORMAT_BACKEND))
    End With
    If mainService.IsEverythingOkayInEngine Then
        Call view.AfterPriceFormCRUDOperation(TypeOfOperation, True)
    End If
End Sub

Private Sub presenterPriceFormFrame_UpdateRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
    Call RefreshMainTable
    Dim response As Variant
    With PFModel
    response = model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_MAINRECORDS, .index, .RecordID)
        If response = True Then
            Call mainService.UpdateRecordOfContextTable(.index, _
                                                        .RecordID, _
                                                        .userID, _
                                                        .recordStatus, _
                                                        VBA.Format$(.statusChangeDate, DATEFORMAT_BACKEND), _
                                                        .conditionType, _
                                                        .salesOrganization, _
                                                        .distributionChannel, _
                                                        .customerID, _
                                                        .materialID, _
                                                        .price, _
                                                        .currencyType, _
                                                        .unitOfPrice, _
                                                        .unitOfMeasure, _
                                                        VBA.Format$(.validFromDate, DATEFORMAT_BACKEND), _
                                                        VBA.Format$(.validToDate, DATEFORMAT_BACKEND))
            If mainService.IsEverythingOkayInEngine Then
                Call view.AfterPriceFormCRUDOperation(TypeOfOperation, True)
            End If
        Else
            Call view.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub presenterPriceFormFrame_DeleteRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
    Call RefreshMainTable
    Dim response As Variant
    With PFModel
    response = model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_MAINRECORDS, .index, .RecordID)
        If response = True Then
            Call mainService.DeleteRecordofContextTable(.index)
            If mainService.IsEverythingOkayInEngine Then
                Call view.AfterPriceFormCRUDOperation(TypeOfOperation, True)
            End If
        Else
            Call view.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub presenterPriceFormFrame_ClosePriceForm()
    Call view.UserWantsToCloseFrame(FORM_PRICEFORM)
End Sub

Private Sub presenterPriceFormFrame_OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
    Call view.ShowWarning(message, typeOfMessage)
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
    response = model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_MAINRECORDS, .index, .RecordID)
        If response = True Then
            Call mainService.UpdateRecordOfContextTable(.index, , _
                                                        .RecordID, _
                                                        .userID, _
                                                        .recordStatus, _
                                                        VBA.Format$(.statusChangeDate, DATEFORMAT_BACKEND), _
                                                        .conditionType, _
                                                        .salesOrganization, _
                                                        .distributionChannel, _
                                                        .customerID, _
                                                        .materialID, _
                                                        .price, _
                                                        .currencyType, _
                                                        .unitOfPrice, _
                                                        .unitOfMeasure, _
                                                        VBA.Format$(.validFromDate, DATEFORMAT_BACKEND), _
                                                        VBA.Format$(.validToDate, DATEFORMAT_BACKEND))
            If mainService.IsEverythingOkayInEngine Then
                Call view.AfterPriceFormCRUDOperation(TypeOfOperation, True)
            End If
        Else
            Call view.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub presenterPriceFormFrame_RejectRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
    Call RefreshMainTable
    Dim response As Variant
    With PFModel
    response = model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_MAINRECORDS, .index, .RecordID)
        If response = True Then
            Call mainService.UpdateRecordOfContextTable(.index, , _
                                                        .RecordID, _
                                                        .userID, _
                                                        .recordStatus, _
                                                        VBA.Format$(.statusChangeDate, DATEFORMAT_BACKEND), _
                                                        .conditionType, _
                                                        .salesOrganization, _
                                                        .distributionChannel, _
                                                        .customerID, _
                                                        .materialID, _
                                                        .price, _
                                                        .currencyType, _
                                                        .unitOfPrice, _
                                                        .unitOfMeasure, _
                                                        VBA.Format$(.validFromDate, DATEFORMAT_BACKEND), _
                                                        VBA.Format$(.validToDate, DATEFORMAT_BACKEND))
            If mainService.IsEverythingOkayInEngine Then
                Call view.AfterPriceFormCRUDOperation(TypeOfOperation, True)
            End If
        Else
            Call view.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

'-------------------------------------------------------------------------
'Data FORM Events
'-------------------------------------------------------------------------

Private Sub presenterDataFormFrame_OpenDataForm(ByVal DFModel As DataFormModel, ByVal ContainerIdentifier As DataContainer)
    'Hydrate Model Properties!
    With DFModel
        Select Case ContainerIdentifier
            Case DataContainer.FOR_CLIENTHISTORY
                .dataTable = mainService.GetSpecificListOfRecordsFromContextTable(MainTableFields.COL_MAIN_userID, model.ActiveUserID)
            Case DataContainer.FOR_PENDINGAPPROVALS
                .dataTable = mainService.GetSpecificListOfRecordsFromContextTable(MainTableFields.COL_MAIN_recordStatus, "PENDING")
            Case DataContainer.FOR_ALLHISTORY
                .dataTable = model.mainTable
        End Select
        .usersTable = model.usersTable
        .ActiveDataContainer = ContainerIdentifier
    End With
    If mainService.IsEverythingOkayInEngine Then
        'Call View Method to Open Data Form Frame!
        Call view.UserWantsToOpenDataFormFrame(DFModel, ContainerIdentifier)
    End If
End Sub

Private Sub presenterDataFormFrame_EditRecordFromDataForm(ByVal DFModel As DataFormModel)
    'We are leaving here one frame and entering into another frame hence,
    'we shold inform main model about the index that next frame should begin with
    model.CurrentUpdateIndex = DFModel.index
    Call OpenPriceForm(OPERATION_UPDATE)
End Sub

Private Sub presenterDataFormFrame_CloseDataForm()
    Call view.UserWantsToCloseFrame(FORM_DATAFORM)
End Sub

Private Sub presenterDataFormFrame_OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
    Call view.ShowWarning(message, typeOfMessage)
End Sub

'-------------------------------------------------------------------------
'Export FORM Events
'-------------------------------------------------------------------------

Private Sub presenterExportFormFrame_OpenExportForm(ByVal EFModel As ExportFormModel)
    'Hydrate Export Frame Model Properties
    With model
        EFModel.InItModel .mainTable, _
                          .usersTable, _
                          .GetUniqueListFromColumn(.mainTable, MainTableFields.COL_MAIN_customerID), _
                          .GetUniqueListFromColumn(.usersTable, UsersTableFields.COL_userName), _
                          .recordStatusList
    End With
    'Open Export Form Frame
    Call view.UserWantsToOpenExportFormFrame(EFModel)
End Sub

Private Sub presenterExportFormFrame_ExportRecordsFromExportForm(ByVal EFModel As ExportFormModel)
    With report
        'Configure Report!
        .CurrencyColumns = Array(MainTableFields.COL_MAIN_price)
        .CURRENCYFORMAT = modGlobal.EXPORTREPORT_CURRENCYFORMAT
        .DateColumns = Array(MainTableFields.COL_MAIN_statusChangeDate, _
                             MainTableFields.COL_MAIN_validFromDate, _
                             MainTableFields.COL_MAIN_validToDate)
        .DateFormat = DATEFORMAT_FRONTEND
        .LoadData = EFModel.exportTable
        .Headers = modDataSources.arrHeaders_Export_Report
        .GenerateReport
        'After Report Generation process
        If .IsReportGeneratedSuccessfully Then
            Call view.AfterExportOperation(True)
        End If
    End With
End Sub

Private Sub presenterExportFormFrame_CloseExportForm()
    Call view.UserWantsToCloseFrame(FORM_EXPORTUTILITY)
End Sub

Private Sub presenterExportFormFrame_ResetExportForm()
    Call OpenExportForm
End Sub

Private Sub presenterExportFormFrame_OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
    Call view.ShowWarning(message, typeOfMessage)
End Sub

'-------------------------------------------------------------------------
'Private Methods to increase abstractiveness of presenter code
'-------------------------------------------------------------------------

Private Sub OpenLoginForm()
    'Refresh Model Data Table with Updated information
    Call RefreshUsersTable '(Only user because Login Mechanism does not required other tables)
    If mainService.IsEverythingOkayInEngine Then
        'Open New Login Form
        If presenterLoginFrame Is Nothing Then Set presenterLoginFrame = New LoginFormPresenter
        'InIt Presenter from here
        Call presenterLoginFrame.InItLoginForm(view)
    End If
End Sub

Private Sub OpenPasswordManager()
    'Refresh Model Data Table with Updated information
    Call RefreshUsersTable '(Only user because password manager does not required other tables)
    If userService.IsEverythingOkayInEngine Then
        'Open Password Manager
        If presenterPasswordManagerFrame Is Nothing Then Set presenterPasswordManagerFrame = New PasswordManagerPresenter
        'InIt Presenter from here
        Call presenterPasswordManagerFrame.InItPasswordManagerForm(view)
    End If
End Sub

Private Sub OpenUserManager()
    'Refresh Model Data Table with Updated information
    Call RefreshUsersTable '(Only user because password manager does not required other tables)
    If userService.IsEverythingOkayInEngine Then
        'open usermanager
        If presenterUserManagerFrame Is Nothing Then Set presenterUserManagerFrame = New UserManagerPresenter
        'InIt Presenter from here
        Call presenterUserManagerFrame.InItUserManagerForm(view)
    End If
End Sub

Private Sub OpenPriceForm(ByVal operation As FormOperation)
    'Refresh Model Data Table with Updated information
    Call RefreshMainTable '(Only user because password manager does not required other tables)
    If mainService.IsEverythingOkayInEngine Then
        'open usermanager
        If presenterPriceFormFrame Is Nothing Then Set presenterPriceFormFrame = New PriceFormPresenter
        If operation = OPERATION_NEW Then
            'InIt Presenter from here
            Call presenterPriceFormFrame.InItPriceForm(view, OPERATION_NEW)
        Else
            'Init Presenter From here for update process
            Call presenterPriceFormFrame.InItPriceForm(view, OPERATION_UPDATE)
        End If
    End If
End Sub

Private Sub OpenDataForm(ByVal ContainerIdentification As DataContainer)
    'Refresh model data table with updated information
    Call RefreshMainTable
    If mainService.IsEverythingOkayInEngine Then
        'Open DataForm
        If presenterDataFormFrame Is Nothing Then Set presenterDataFormFrame = New DataFormPresenter
        'InIt Presenter from here
        Call presenterDataFormFrame.InItDataForm(view, ContainerIdentification)
    End If
End Sub

Private Sub OpenExportForm()
    'Refresh Model Data Table with Updated information
    Call RefreshMainTable '(Only user because password manager does not required other tables)
    If mainService.IsEverythingOkayInEngine Then
        'open usermanager
        If presenterExportFormFrame Is Nothing Then Set presenterExportFormFrame = New ExportFormPresenter
        'InIt Presenter from here
        Call presenterExportFormFrame.InItExportForm(view)
    End If
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

Private Sub Class_Initialize()
    Set report = New DocumentGeneratorExcel
End Sub

Private Sub Class_Terminate()
    Set presenterExportFormFrame = Nothing
    Set presenterDataFormFrame = Nothing
    Set presenterPriceFormFrame = Nothing
    Set presenterUserManagerFrame = Nothing
    Set presenterPasswordManagerFrame = Nothing
    Set presenterLoginFrame = Nothing
    Set view = Nothing
    Set model = Nothing
    Set userService = Nothing
    Set mainService = Nothing
    Set report = Nothing
End Sub



