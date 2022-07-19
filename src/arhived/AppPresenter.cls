VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@Folder("MVP.Presenter")
'@PredeclaredId
Option Explicit

Implements IAppPresenter
Implements IDisposable

'-------------------------------------------------------------------------
'Event handllers
'-------------------------------------------------------------------------

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents View As PriceApprovalView
Attribute View.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents PresenterLoginFrame As LoginFormPresenter
Attribute PresenterLoginFrame.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents PresenterPasswordManagerFrame As PasswordManagerPresenter
Attribute PresenterPasswordManagerFrame.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents PresenterUserManagerFrame As UserManagerPresenter
Attribute PresenterUserManagerFrame.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents PresenterPriceFormFrame As PriceFormPresenter
Attribute PresenterPriceFormFrame.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents PresenterDataFormFrame As DataFormPresenter
Attribute PresenterDataFormFrame.VB_VarHelpID = -1
'@MemberAttribute VB_VarHelpID, -1
Private WithEvents PresenterExportFormFrame As ExportFormPresenter
Attribute PresenterExportFormFrame.VB_VarHelpID = -1

'-------------------------------------------------------------------------
'Private Componenets Set
'-------------------------------------------------------------------------

Private Type TAppPresenterComponents
    context As IAppContext
    Model As AppModel
    MainView As IView
    Report As DocumentGeneratorExcel
    Disposed As Boolean
End Type

Private this As TAppPresenterComponents

'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As AppPresenter
Attribute Class.VB_Description = "Returns class reference"
    Set Class = AppPresenter
End Property

'-------------------------------------------------------------------------
'Private properties
'-------------------------------------------------------------------------

Public Property Set MainView(ByVal RHS As IView)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization this.MainView
    Guard.NullReference RHS
    Set this.MainView = RHS
    Set View = RHS
End Property

'-------------------------------------------------------------------------

Public Property Get Model() As AppModel
    Guard.DefaultInstance Me
    Set Model = this.Model
End Property

Public Property Set Model(ByVal RHS As AppModel)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization this.Model
    Guard.NullReference RHS
    Set this.Model = RHS
End Property

Public Property Get context() As IAppContext
    Guard.DefaultInstance Me
    Set context = this.context
End Property

Public Property Set context(ByVal RHS As IAppContext)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization this.context
    Guard.NullReference RHS
    Set this.context = RHS
End Property

Public Property Get Report() As DocumentGeneratorExcel
    Guard.DefaultInstance Me
    Set Report = this.Report
End Property

Public Property Set Report(ByVal RHS As DocumentGeneratorExcel)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization this.Report
    Guard.NullReference RHS
    Set this.Report = RHS
End Property


Public Function Create(ByVal context As IAppContext, ByVal Model As AppModel, ByVal View As IView) As IAppPresenter
    Guard.NonDefaultInstance Me
    Guard.NullReference context
    Guard.NullReference Model
    Guard.NullReference View

    Dim result As AppPresenter
    Set result = New AppPresenter
    
    Set result.context = context
    Set result.MainView = View
    Set result.Model = Model
    Set result.Report = New DocumentGeneratorExcel
    Set Create = result

End Function

'-------------------------------------------------------------------------
'Frame Opening Operations / Side Panel Buttons
'-------------------------------------------------------------------------

Private Sub View_OpenLoginFrame()
    Call OpenLoginForm
End Sub

Private Sub View_OpenPasswordManagerFrame()
    Call OpenPasswordManager
End Sub

Private Sub View_OpenUserManagerFrame()
    Call OpenUserManager
End Sub

Private Sub View_OpenPriceFormFrame()
    Call OpenPriceForm(OPERATION_NEW)
End Sub

Private Sub View_OpenDataFormFrame(ByVal ContainerIdentifier As DataContainer)
    Select Case ContainerIdentifier
    Case DataContainer.FOR_CLIENTHISTORY
        Call OpenDataForm(FOR_CLIENTHISTORY)
    Case DataContainer.FOR_PENDINGAPPROVALS
        Call OpenDataForm(FOR_PENDINGAPPROVALS)
    Case DataContainer.FOR_ALLHISTORY
        Call OpenDataForm(FOR_ALLHISTORY)
    End Select
End Sub

Private Sub View_OpenExportFormFrame()
    Call OpenExportForm
End Sub

Private Sub View_Logout()
    Call View.UserWantsToLogout
End Sub

Private Sub View_ExitApp()
    Call View.OnCancel
End Sub

'-------------------------------------------------------------------------
'Frame Operations
'-------------------------------------------------------------------------


'------------------
'Login Frame Events
'------------------

Private Sub PresenterLoginFrame_OpenLoginForm(ByVal LModel As LoginFormModel)
    'Hydrate model property
    Call LModel.InItModel(Model.usersTable)
    'Open Login Form
    Call View.UserWantsToOpenLoginFrame(LModel)
End Sub

Private Sub PresenterLoginFrame_Login()
    Call View.UserWantsToLogin
End Sub

Private Sub PresenterLoginFrame_CancelLoginForm()
    Call View.UserWantsToCloseFrame(FORM_LOGIN)
End Sub

Private Sub PresenterLoginFrame_ResetPasswordForUser(ByVal targetUserName As String, ByVal TargetEmailAddress As String)
    
    Dim dummyPassword   As String
    Dim NewPassword     As String
    Dim TargetIndex     As Long
    Dim TargetRow       As Long
    Dim MessageToShow   As String
    
    'Generate Dummy Password
    dummyPassword = AppMethods.RandomString(10)
    NewPassword = dummyPassword
    
    'Reset User's Password
    TargetIndex = AppMethods.GetLookupValue(Model.usersTable, _
                                            targetUserName, _
                                            UsersTableFields.COL_userName, _
                                            UsersTableFields.COL_INDEX)
                                            
    TargetRow = AppMethods.GetTargetRowIndex(Model.usersTable, _
                                             TargetIndex, _
                                             UsersTableFields.COL_INDEX)
                                            
    '.userIndex, .userID, .userStatus, .UserType, .UserName, .userPassword, .userEmail
    Call context.UserService.UpdateRecordOfContextTable(TargetIndex, _
        Model.usersTable(TargetRow, UsersTableFields.COL_userId), _
        Model.usersTable(TargetRow, UsersTableFields.COL_userStatus), _
        Model.usersTable(TargetRow, UsersTableFields.COL_userType), _
        Model.usersTable(TargetRow, UsersTableFields.COL_userName), _
        dummyPassword, _
        Model.usersTable(TargetRow, UsersTableFields.COL_email))
    
    If context.UserService.IsEverythingOkayInEngine Then
    
        MessageToShow = "Dear " & Model.usersTable(TargetRow, UsersTableFields.COL_userName) & "," & vbNewLine & vbNewLine & _
                        "Your new password is following" & vbNewLine & vbNewLine & _
                        NewPassword & vbNewLine & vbNewLine & _
                        "Kindly note down the password and Login again with new password."
        
        'Notify user through messagebox that, new password has been sent to you email.
        Call View.ShowWarning(MessageToShow, TYPE_INFORMATION)
        
    Else
        Call View.ShowWarning("Something went wrong while Resetting the password. Please try again", TYPE_CRITICAL)
    End If
    
    Call OpenLoginForm
    
End Sub

Private Sub PresenterLoginFrame_OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
    Call View.ShowWarning(message, typeOfMessage)
End Sub

'------------------------
'Password Manager Events
'------------------------

Private Sub PresenterPasswordManagerFrame_OpenPasswordManagerForm(ByVal PMModel As PasswordManagerModel)
    'Hydrate Model property
    Call PMModel.InItModel(Model.ActiveUserPassword)
    'Open Passwordmanager Frame
    Call View.UserWantsToOpenPasswordManagerFrame(PMModel)
End Sub

Private Sub PresenterPasswordManagerFrame_ChangePassword(ByVal NewPassword As String)
    'index, userID, userStatus, userType, userName, password
    With Model
        Call context.UserService.UpdateRecordOfContextTable(.ActiveUserIndex, .ActiveUserID, .ActiveUserStatus, .ActiveUserType, .ActiveUserName, NewPassword, .ActiveUserEmail)
        If context.UserService.IsEverythingOkayInEngine Then
            Call View.AfterChangePasswordOperation
        End If
    End With
End Sub

Private Sub PresenterPasswordManagerFrame_OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
    Call View.ShowWarning(message, typeOfMessage)
End Sub

Private Sub PresenterPasswordManagerFrame_CancelPasswordManagerForm()
    Call View.UserWantsToCloseFrame(FORM_PASSWORDMANAGER)
End Sub

'-------------------
'User Manager Events
'-------------------

Private Sub PresenterUserManagerFrame_OpenUserManagerForm(ByVal UMModel As UserManagerModel)
    'Hydrate Model Properties
    Call UMModel.InItModel(Model.usersTable, Model.userStatusList, Model.userTypeList)
    'Open User manager frame
    Call View.UserWantsToOpenUserManagerFrame(UMModel)
End Sub

Private Sub PresenterUserManagerFrame_CloseUserManagerForm()
    Call View.UserWantsToCloseFrame(FORM_USERMANAGER)
End Sub

Private Sub PresenterUserManagerFrame_ResetUserManagerForm()
    Call OpenUserManager
End Sub

Private Sub PresenterUserManagerFrame_UpdateUserManagerFormRecord()
    Call View.UserWantsToUpdateUserManagerRecord
End Sub

'Values shouold be in sequence of how they are in table!
'Columns of Users Table
'0. index
'1. userID
'2. userStatus
'3. userType
'4. userName
'5. password
'6. Email
'index will be alloted automatically so we should start from 1 to 5

Private Sub PresenterUserManagerFrame_AddNewUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal UMModel As UserManagerModel)
    Call RefreshUsersTable
    With UMModel
        
        context.UserService.AddNewRecordToContextTable Model.GetNewID(TablesOfThisApplication.TABLE_USERS), _
        .userStatus, _
        .userType, _
        .UserName, _
        .userPassword, _
        .userEmail
        
        Dim msgPrompt As String
        msgPrompt = "New User Added sucessfully." & VBA.vbNewLine & _
                    "Notify user: " & .UserName & "?"
        
        NotifyUser msgPrompt, OPERATION_NEW, .UserName, .userEmail, .userPassword
        
    End With
    If context.UserService.IsEverythingOkayInEngine Then
        Call View.AfterUserManagerCRUDOperation(TypeOfOperation, True)
    End If
End Sub

Private Sub PresenterUserManagerFrame_UpdateUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal UMModel As UserManagerModel)
    Call RefreshUsersTable
    Dim response As Variant
    With UMModel
        response = Model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_USERS, .userIndex, .userID)
        If response = True Then
            context.UserService.UpdateRecordOfContextTable .userIndex, _
                .userID, _
                .userStatus, _
                .userType, _
                .UserName, _
                .userPassword, _
                .userEmail
                
            Dim msgPrompt As String
            msgPrompt = "User Updated sucessfully." & VBA.vbNewLine & _
                    "Notify user: " & .UserName & "?"
        
            NotifyUser msgPrompt, OPERATION_UPDATE, .UserName, .userEmail, .userPassword
                
            If context.UserService.IsEverythingOkayInEngine Then
                Call View.AfterUserManagerCRUDOperation(TypeOfOperation, True)
            End If
        Else
            Call View.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub PresenterUserManagerFrame_DeleteUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal UMModel As UserManagerModel)
    Call RefreshUsersTable
    Dim response As Variant
    With UMModel
        response = Model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_USERS, .userIndex, .userID)
        If response = True Then
            Call context.UserService.DeleteRecordofContextTable(.userIndex)
            
            Dim msgPrompt As String
            msgPrompt = "User Deleted sucessfully." & VBA.vbNewLine & _
                    "Notify user: " & .UserName & "?"
        
            NotifyUser msgPrompt, OPERATION_DELETE, .UserName, .userEmail
            
            If context.UserService.IsEverythingOkayInEngine Then
                Call View.AfterUserManagerCRUDOperation(TypeOfOperation, True)
            End If
        Else
            Call View.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub NotifyUser(ByVal msgPrompt As String, ByVal opType As FormOperation, ByVal uName As String, ByVal uMail As String, _
                        Optional ByVal uPassword As String)
    
    Dim userAnswer As Long
    userAnswer = AppMethods.UserNotifierConfirmation(msgPrompt, "Please confirm")
                                        
    If Not userAnswer = VBA.vbNo Then
    
        Select Case opType
            Case Is = FormOperation.OPERATION_NEW
                EmailServices.EmailNotifyUser uMail, _
                                            "User registered in the Price Approval App", _
                                            False, _
                                            "See below your login credentials: <B></B><br><br>", _
                                            "<B>Username: </B></td><td>" & uName, _
                                            "<B>Password: </B></td><td>" & uPassword, _
                                            "<B><em>Consider changing the password at your first login. </em></B>"
            
            Case Is = FormOperation.OPERATION_UPDATE
                EmailServices.EmailNotifyUser uMail, _
                                            "User updated in the Price Approval App", _
                                            False, _
                                            "See below your new login credentials: <B></B><br><br>", _
                                            "<B>Username: </B></td><td>" & uName, _
                                            "<B>Password: </B></td><td>" & uPassword, _
                                            "<B><em>Consider changing the password at your next login. </em></B>"
            
            Case Is = FormOperation.OPERATION_DELETE
                EmailServices.EmailNotifyUser uMail, _
                                            "User removed from the Price Approval App", _
                                            False, _
                                            "This user was removed from the application: <B></B><br><br>", _
                                            "<B>Username: </B></td><td>" & uName, _
                                            "<B>" & vbNullString & "</B></td><td>" & vbNullString, _
                                            "<B><em>" & vbNullString & "</em></B>"
        
        End Select
    
    End If

End Sub

Private Sub PresenterUserManagerFrame_OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
    Call View.ShowWarning(message, typeOfMessage)
End Sub

'-------------------
'PRICE FORM Events
'-------------------

Private Sub PresenterPriceFormFrame_OpenPriceForm(ByVal PFModel As PriceFormModel, ByVal operation As FormOperation)
    If operation = OPERATION_NEW Then
        'Hydrate Model property
        With Model
            Call PFModel.InItModel(.mainTable, .CurrenciesList, .unitOfMeasuresList, .salesOrganizationList, .distributionChannelsList)
        End With
        'Open PriceForm Frame
        Call View.UserWantsToOpenPriceFormFrame(PFModel, OPERATION_NEW)
    Else
        With Model
            Call PFModel.InItModel(.mainTable, .CurrenciesList, .unitOfMeasuresList, .salesOrganizationList, .distributionChannelsList, .CurrentUpdateIndex)
        End With
        'Open PriceForm Frame
        Call View.UserWantsToOpenPriceFormFrame(PFModel, OPERATION_UPDATE)
    End If
End Sub

Private Sub PresenterPriceFormFrame_ResetPriceForm()
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

Private Sub PresenterPriceFormFrame_AddNewRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
    Call RefreshMainTable
    'I added here StatusChangeDate, because, While filtering the date it is causing an issue! Consider this statusChangeDate field as record added/update date!
    With PFModel
        Call context.MainService.AddNewRecordToContextTable(Model.GetNewID(TablesOfThisApplication.TABLE_MAINRECORDS), _
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
    If context.MainService.IsEverythingOkayInEngine Then
        Call View.AfterPriceFormCRUDOperation(TypeOfOperation, True)
    End If
End Sub

Private Sub PresenterPriceFormFrame_UpdateRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
    Call RefreshMainTable
    Dim response As Variant
    With PFModel
        response = Model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_MAINRECORDS, .Index, .RecordID)
        If response = True Then
            Call context.MainService.UpdateRecordOfContextTable(.Index, _
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
            If context.MainService.IsEverythingOkayInEngine Then
                Call View.AfterPriceFormCRUDOperation(TypeOfOperation, True)
            End If
        Else
            Call View.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub PresenterPriceFormFrame_DeleteRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
    Call RefreshMainTable
    Dim response As Variant
    With PFModel
        response = Model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_MAINRECORDS, .Index, .RecordID)
        If response = True Then
            Call context.MainService.DeleteRecordofContextTable(.Index)
            If context.MainService.IsEverythingOkayInEngine Then
                Call View.AfterPriceFormCRUDOperation(TypeOfOperation, True)
            End If
        Else
            Call View.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub PresenterPriceFormFrame_ClosePriceForm()
    Call View.UserWantsToCloseFrame(FORM_PRICEFORM)
End Sub

Private Sub PresenterPriceFormFrame_OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
    Call View.ShowWarning(message, typeOfMessage)
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

Private Sub PresenterPriceFormFrame_ApproveRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
    Call RefreshMainTable
    Dim response As Variant
    With PFModel
        response = Model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_MAINRECORDS, .Index, .RecordID)
        If response = True Then
            Call context.MainService.UpdateRecordOfContextTable(.Index, _
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
            If context.MainService.IsEverythingOkayInEngine Then
                Call View.AfterPriceFormCRUDOperation(TypeOfOperation, True)
            End If
        Else
            Call View.ShowWarning(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub PresenterPriceFormFrame_RejectRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
    
    Call RefreshMainTable
    
    Dim response    As Variant
    Dim uName       As String
    Dim uMail       As String
    Dim subject     As String
    
    With PFModel
        response = Model.IsRecordIDAvailableToUpdate(TablesOfThisApplication.TABLE_MAINRECORDS, .Index, .RecordID)
        
        If response = True Then
        
            Call context.MainService.UpdateRecordOfContextTable(.Index, _
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
            'EMail Creation Starts
            'Sending Rejection Email to the Client
            uName = AppMethods.GetLookupValue(Model.usersTable, .userID, UsersTableFields.COL_userId, UsersTableFields.COL_userName)
            uMail = AppMethods.GetLookupValue(Model.usersTable, .userID, UsersTableFields.COL_userId, UsersTableFields.COL_email)
            subject = "Price Approval: Record: #" & .RecordID & "#" & " has been Rejected."

            'Confirm that approver wants to notify the client user
            Dim msgPrompt As String: msgPrompt = "Notify user: " & uName & "?"
            Dim userAnswer As Long
            userAnswer = AppMethods.UserNotifierConfirmation(msgPrompt, "Please confirm")
                                                
            If userAnswer = VBA.vbNo Then        'Approver cancelled process to notify
                GoTo UserCancelled
                
            Else
                                                
                Dim targetMessage As String
                If ValidationServices.IsInputValid(outText:=targetMessage, inPrompt:="Rejecting reason: ", inTitel:="Please type meassage") Then
                   EmailServices.EmailRecordRejected uName, _
                                         uMail, _
                                         subject, _
                                         False, _
                                         "Message from approver: <B>" & targetMessage & "</B><br><br>" & _
                                         "Details of the record that has been rejected:<br>", _
                                         "<B>Record </B></td><td>" & .RecordID, _
                                         "<B>UserID </B></td><td>" & .userID, _
                                         "<B>Staus of Record </B></td><td>" & .recordStatus, _
                                         "<B>Condition Type </B></td><td>" & .conditionType, _
                                         "<B>Sales Organization </B></td><td>" & .salesOrganization, _
                                         "<B>Distribution Channel </B></td><td>" & .distributionChannel, _
                                         "<B>Customer ID </B></td><td>" & .customerID, _
                                         "<B>Material ID </B></td><td>" & .materialID, _
                                         "<B>Price Proposed </B></td><td>" & .price, _
                                         "<B>Currency </B></td><td>" & .currencyType, _
                                         "<B>Unit of Price </B></td><td>" & .unitOfPrice, _
                                         "<B>Unit of Measure </B></td><td>" & .unitOfMeasure, _
                                         "<B>Valid from Date </B></td><td>" & VBA.Format$(.validFromDate, GetDateFormat), _
                                         "<B>Valid To Date </B></td><td>" & VBA.Format$(.validToDate, GetDateFormat)
                End If 'EMail Creation Ends
                
            End If
            
UserCancelled:

            If context.MainService.IsEverythingOkayInEngine Then
                Call View.AfterPriceFormCRUDOperation(TypeOfOperation, True)
            End If
        Else
            Call View.ShowWarning(response, TYPE_CRITICAL)
        End If
        
    End With
    
End Sub

'-------------------------------------------------------------------------
'Data FORM Events
'-------------------------------------------------------------------------

Private Sub PresenterDataFormFrame_OpenDataForm(ByVal DFModel As DataFormModel, ByVal ContainerIdentifier As DataContainer)
    'Hydrate Model Properties!
    With DFModel
        Select Case ContainerIdentifier
        Case DataContainer.FOR_CLIENTHISTORY
            .dataTable = context.MainService.GetSpecificListOfRecordsFromContextTable(MainTableFields.COL_MAIN_userID, Model.ActiveUserID)
        Case DataContainer.FOR_PENDINGAPPROVALS
            .dataTable = context.MainService.GetSpecificListOfRecordsFromContextTable(MainTableFields.COL_MAIN_recordStatus, "PENDING")
        Case DataContainer.FOR_ALLHISTORY
            .dataTable = Model.mainTable
        End Select
        .usersTable = Model.usersTable
        .ActiveDataContainer = ContainerIdentifier
        .DataColumnsList = DataResources.arrListOfColumnsMainTable
    End With
    
    If context.MainService.IsEverythingOkayInEngine Then
        'Call View Method to Open Data Form Frame!
        Call View.UserWantsToOpenDataFormFrame(DFModel, ContainerIdentifier)
    End If
    
End Sub

Private Sub PresenterDataFormFrame_ResetDataForm(ByVal DFModel As DataFormModel, ByVal ContainerIdentifier As DataContainer)
    'Hydrate Model Properties!
    With DFModel
        Select Case ContainerIdentifier
        Case DataContainer.FOR_CLIENTHISTORY
            .dataTable = context.MainService.GetSpecificListOfRecordsFromContextTable(MainTableFields.COL_MAIN_userID, Model.ActiveUserID)
        Case DataContainer.FOR_PENDINGAPPROVALS
            .dataTable = context.MainService.GetSpecificListOfRecordsFromContextTable(MainTableFields.COL_MAIN_recordStatus, "PENDING")
        Case DataContainer.FOR_ALLHISTORY
            .dataTable = Model.mainTable
        End Select
        .usersTable = Model.usersTable
        .ActiveDataContainer = ContainerIdentifier
        .DataColumnsList = DataResources.arrListOfColumnsMainTable
    End With
    
    If context.MainService.IsEverythingOkayInEngine Then
        'Call View Method to Open Data Form Frame!
        Call View.UserWantsToOpenDataFormFrame(DFModel, ContainerIdentifier)
    End If
    
End Sub

Private Sub PresenterDataFormFrame_FilterAndSortDataFormList()
    Call View.UserWantsToFilterAndSortDataFormList
End Sub

Private Sub PresenterDataFormFrame_PopulateValuesComboBox(ByVal DFModel As DataFormModel, ByVal TargetColumnIndex As Long)
    DFModel.ValuesList = Model.GetUniqueListFromColumn(DFModel.dataTable, TargetColumnIndex)
    Call View.ApplicationWantsToUpdateValueListComboBox
End Sub

Private Sub PresenterDataFormFrame_EditRecordFromDataForm(ByVal DFModel As DataFormModel)
    'We are leaving here one frame and entering into another frame hence,
    'we shold inform main model about the index that next frame should begin with
    Model.CurrentUpdateIndex = DFModel.Index
    Call OpenPriceForm(OPERATION_UPDATE)
End Sub

Private Sub PresenterDataFormFrame_CloseDataForm()
    Call View.UserWantsToCloseFrame(FORM_DATAFORM)
End Sub

Private Sub PresenterDataFormFrame_OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
    Call View.ShowWarning(message, typeOfMessage)
End Sub

'-------------------------------------------------------------------------
'Export FORM Events
'-------------------------------------------------------------------------

Private Sub PresenterExportFormFrame_OpenExportForm(ByVal EFModel As ExportFormModel)
    'Hydrate Export Frame Model Properties
    With Model
        EFModel.InItModel .mainTable, _
                          .usersTable, _
                          .GetUniqueListFromColumn(.mainTable, MainTableFields.COL_MAIN_customerID), _
                          .GetUniqueListFromColumn(.usersTable, UsersTableFields.COL_userName), _
                          .recordStatusList
    End With
    'Open Export Form Frame
    Call View.UserWantsToOpenExportFormFrame(EFModel)
End Sub

Private Sub PresenterExportFormFrame_ExportRecordsFromExportForm(ByVal EFModel As ExportFormModel)
    With Report
    
        'Configure Report!
        .CurrencyColumns = Array(MainTableFields.COL_MAIN_price)
        .CURRENCYFORMAT = GlobalResources.EXPORTREPORT_CURRENCYFORMAT
        .DateColumns = Array(MainTableFields.COL_MAIN_statusChangeDate, _
                             MainTableFields.COL_MAIN_validFromDate, _
                             MainTableFields.COL_MAIN_validToDate)
        
        'Process the Data Before we produce the report in excel format!
        If EFModel.IsRecordsNeedsToBeProcessed Then
            
            Dim Records()       As String
            Dim Record          As Variant
            Dim TargetRowIndex  As Long
            Dim TotalRecords    As Long
            Dim counter         As Long
            
            Records = VBA.Split(EFModel.RecordsToProcess, SEPERATOR_ITEM)
            TotalRecords = UBound(Records) + 1
                        
            View.ShowStatusOfExportProcess "Process start..."
            VBA.DoEvents
            counter = 0
            
            With Model
                For Each Record In Records
                    context.MainService.UpdateRecordOfContextTable .mainTable(Record, 1), _
        .mainTable(Record, 2), _
        .mainTable(Record, 3), _
        GlobalResources.RECORDSTATUS_PROCESSED, _
        VBA.Format$(VBA.Now, DATEFORMAT_BACKEND), _
        .mainTable(Record, 6), _
        .mainTable(Record, 7), _
        .mainTable(Record, 8), _
        .mainTable(Record, 9), _
        .mainTable(Record, 10), _
        .mainTable(Record, 11), _
        .mainTable(Record, 12), _
        .mainTable(Record, 13), _
        .mainTable(Record, 14), _
        VBA.Format$(.mainTable(Record, 15), DATEFORMAT_BACKEND), _
        VBA.Format$(.mainTable(Record, 16), DATEFORMAT_BACKEND)
                    counter = counter + 1
                    View.ShowStatusOfExportProcess counter & " of " & TotalRecords & " Records has been processed!"
                    VBA.DoEvents
                Next Record
            End With
            
            'Splash screen output
            View.ShowStatusOfExportProcess "All records are processed! Exporting the report now."
            VBA.DoEvents
            
        End If
        
        'Export Excel based report
        .DateFormat = GetDateFormat
        .LoadData = EFModel.exportTable
        .Headers = DataResources.arrHeaders_Export_Report
        .GenerateReport
        
        'hide status
        View.ShowStatusOfExportProcess vbNullString
        
        'After Report Generation process
        If .IsReportGeneratedSuccessfully Then
            Call View.AfterExportOperation(True)
        End If
        
    End With
End Sub

Private Sub PresenterExportFormFrame_CloseExportForm()
    Call View.UserWantsToCloseFrame(FORM_EXPORTUTILITY)
End Sub

Private Sub PresenterExportFormFrame_ResetExportForm()
    Call OpenExportForm
End Sub

Private Sub PresenterExportFormFrame_OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
    Call View.ShowWarning(message, typeOfMessage)
End Sub

'-------------------------------------------------------------------------
'Private Methods to increase abstractiveness of presenter code
'-------------------------------------------------------------------------

Private Sub OpenLoginForm()
    'Refresh Model Data Table with Updated information
    Call RefreshUsersTable                       '(Only user because Login Mechanism does not required other tables)
    If context.MainService.IsEverythingOkayInEngine Then
        'Open New Login Form
        If PresenterLoginFrame Is Nothing Then Set PresenterLoginFrame = New LoginFormPresenter
        'InIt Presenter from here
        Call PresenterLoginFrame.InItLoginForm(View)
    End If
End Sub

Private Sub OpenPasswordManager()
    'Refresh Model Data Table with Updated information
    Call RefreshUsersTable                       '(Only user because password manager does not required other tables)
    If context.UserService.IsEverythingOkayInEngine Then
        'Open Password Manager
        If PresenterPasswordManagerFrame Is Nothing Then Set PresenterPasswordManagerFrame = New PasswordManagerPresenter
        'InIt Presenter from here
        Call PresenterPasswordManagerFrame.InItPasswordManagerForm(View)
    End If
End Sub

Private Sub OpenUserManager()
    'Refresh Model Data Table with Updated information
    Call RefreshUsersTable                       '(Only user because password manager does not required other tables)
    If context.UserService.IsEverythingOkayInEngine Then
        'open usermanager
        If PresenterUserManagerFrame Is Nothing Then Set PresenterUserManagerFrame = New UserManagerPresenter
        'InIt Presenter from here
        Call PresenterUserManagerFrame.InItUserManagerForm(View)
    End If
End Sub

Private Sub OpenPriceForm(ByVal operation As FormOperation)
    'Refresh Model Data Table with Updated information
    Call RefreshMainTable                        '(Only user because password manager does not required other tables)
    If context.MainService.IsEverythingOkayInEngine Then
        'open usermanager
        If PresenterPriceFormFrame Is Nothing Then Set PresenterPriceFormFrame = New PriceFormPresenter
        If operation = OPERATION_NEW Then
            'InIt Presenter from here
            Call PresenterPriceFormFrame.InItPriceForm(View, OPERATION_NEW)
        Else
            'Init Presenter From here for update process
            Call PresenterPriceFormFrame.InItPriceForm(View, OPERATION_UPDATE)
        End If
    End If
End Sub

Private Sub OpenDataForm(ByVal ContainerIdentification As DataContainer)
    'Refresh model data table with updated information
    Call RefreshMainTable
    If context.MainService.IsEverythingOkayInEngine Then
        'Open DataForm
        If PresenterDataFormFrame Is Nothing Then Set PresenterDataFormFrame = New DataFormPresenter
        'InIt Presenter from here
        Call PresenterDataFormFrame.InItDataForm(View, ContainerIdentification)
    End If
End Sub

Private Sub OpenExportForm()
    'Refresh Model Data Table with Updated information
    Call RefreshMainTable                        '(Only user because password manager does not required other tables)
    If context.MainService.IsEverythingOkayInEngine Then
        'open usermanager
        If PresenterExportFormFrame Is Nothing Then Set PresenterExportFormFrame = New ExportFormPresenter
        'InIt Presenter from here
        Call PresenterExportFormFrame.InItExportForm(View)
    End If
End Sub

Private Sub RefreshMainTable()
    Model.mainTable = context.MainService.GetListOfAllRecordsFromContextTable
End Sub

Private Sub RefreshUsersTable()
    Model.usersTable = context.UserService.GetListOfAllRecordsFromContextTable
End Sub

'-------------------------------------------------------------------------
'Class Events
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    If Not this.Disposed Then Dispose
End Sub

Private Sub Dispose()

    If this.Disposed Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If
    
    If Not this.Model Is Nothing Then
        Disposable.TryDispose this.Model
        Set this.Model = Nothing
    End If

    If Not this.context Is Nothing Then
        Disposable.TryDispose this.context
        Set this.context = Nothing
    End If

    If Not this.Report Is Nothing Then
        Disposable.TryDispose this.Report
        Set this.Report = Nothing
    End If

    If Not this.MainView Is Nothing Then
        Disposable.TryDispose this.MainView
        Set this.MainView = Nothing

        Set PresenterExportFormFrame = Nothing
        Set PresenterDataFormFrame = Nothing
        Set PresenterPriceFormFrame = Nothing
        Set PresenterUserManagerFrame = Nothing
        Set PresenterPasswordManagerFrame = Nothing
        Set PresenterLoginFrame = Nothing
        Set View = Nothing
        
    End If
    
    this.Disposed = True
    
    #If TestMode Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " is terminating"
    #End If
    
End Sub

Private Sub IDisposable_Dispose()
    Dispose
End Sub

Private Sub IAppPresenter_ShowView()
    this.MainView.Show
End Sub
