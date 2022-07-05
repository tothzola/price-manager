VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@PredeclaredId
'@Folder("AppObjects.Model")
Option Explicit

'-------------------------------------------------------------------------
'Private Type Componenets / Enums
'-------------------------------------------------------------------------

Private Type TAppModelComponents
  'Data Sources
    mainTable As Variant
    usersTable As Variant
    CurrenciesList As Variant
    unitOfMeasuresList As Variant
    userTypeList As Variant
    userStatusList As Variant
    recordStatusList As Variant
    salesOrganizationList As Variant
    distributionChannelsList As Variant
  'Active User Info
    ActiveUserID As String
    ActiveUserStatus As String
    ActiveUserType As String
    ActiveUserName As String
    ActiveUserPassword As String
    ActiveUserEmail As String
  'Current Parameters
    CurrentUpdateIndex As Long
    
End Type

Private This As TAppModelComponents

'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As AppModel
Attribute Class.VB_Description = "Returns class reference"
    Set Class = AppModel
End Property

'@Description "Creates a new application context."
Public Function Create(ByVal context As IAppContext) As AppModel
Attribute Create.VB_Description = "Creates a new application context."
    Guard.NonDefaultInstance AppModel
    Guard.NullReference context
    
    Dim result As AppModel
    Set result = New AppModel
    
    InitializeAppModel result, context
    
    Set Create = result
End Function

Private Sub InitializeAppModel(ByVal Model As AppModel, ByVal context As IAppContext)
                    
    'Tables of The Applications
    Model.mainTable = context.MainService.GetListOfAllRecordsFromContextTable
    Model.usersTable = context.UserService.GetListOfAllRecordsFromContextTable
  
    'Single Column Lists for DropDown/ComboBoxes
    Model.CurrenciesList = DataResources.arrListofCurrencies
    Model.unitOfMeasuresList = DataResources.arrListOfUnitOfMeasure
    Model.userTypeList = DataResources.arrListofTypesOfUser
    Model.userStatusList = DataResources.arrListofStatusOfUser
    Model.recordStatusList = DataResources.arrRecordStatusesList
    Model.salesOrganizationList = DataResources.arrSalesOrganizationsList
    Model.distributionChannelsList = DataResources.arrDistributionChannelsList
    
End Sub
'-------------------------------------------------------------------------
'Properties
'-------------------------------------------------------------------------

Public Property Get ActiveUserEmail() As String
    ActiveUserEmail = This.ActiveUserEmail
End Property

Public Property Let ActiveUserEmail(ByVal vNewValue As String)
    This.ActiveUserEmail = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ActiveUserPassword() As String
    ActiveUserPassword = This.ActiveUserPassword
End Property

Public Property Let ActiveUserPassword(ByVal vNewValue As String)
    This.ActiveUserPassword = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ActiveUserName() As String
    ActiveUserName = This.ActiveUserName
End Property

Public Property Let ActiveUserName(ByVal vNewValue As String)
    This.ActiveUserName = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ActiveUserType() As String
    ActiveUserType = This.ActiveUserType
End Property

Public Property Let ActiveUserType(ByVal vNewValue As String)
    This.ActiveUserType = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ActiveUserStatus() As String
    ActiveUserStatus = This.ActiveUserStatus
End Property

Public Property Let ActiveUserStatus(ByVal vNewValue As String)
    This.ActiveUserStatus = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ActiveUserID() As String
    ActiveUserID = This.ActiveUserID
End Property

Public Property Let ActiveUserID(ByVal vNewValue As String)
    This.ActiveUserID = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get mainTable() As Variant
    mainTable = This.mainTable
End Property

Public Property Let mainTable(ByVal vNewValue As Variant)
    This.mainTable = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get usersTable() As Variant
    usersTable = This.usersTable
End Property

Public Property Let usersTable(ByVal vNewValue As Variant)
    This.usersTable = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get CurrenciesList() As Variant
    CurrenciesList = This.CurrenciesList
End Property

Public Property Let CurrenciesList(ByVal vNewValue As Variant)
    This.CurrenciesList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get unitOfMeasuresList() As Variant
    unitOfMeasuresList = This.unitOfMeasuresList
End Property

Public Property Let unitOfMeasuresList(ByVal vNewValue As Variant)
    This.unitOfMeasuresList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userTypeList() As Variant
    userTypeList = This.userTypeList
End Property

Public Property Let userTypeList(ByVal vNewValue As Variant)
    This.userTypeList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userStatusList() As Variant
    userStatusList = This.userStatusList
End Property

Public Property Let userStatusList(ByVal vNewValue As Variant)
    This.userStatusList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get recordStatusList() As Variant
    recordStatusList = This.recordStatusList
End Property

Public Property Let recordStatusList(ByVal vNewValue As Variant)
    This.recordStatusList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get distributionChannelsList() As Variant
    distributionChannelsList = This.distributionChannelsList
End Property

Public Property Let distributionChannelsList(ByVal vNewValue As Variant)
    This.distributionChannelsList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get salesOrganizationList() As Variant
    salesOrganizationList = This.salesOrganizationList
End Property

Public Property Let salesOrganizationList(ByVal vNewValue As Variant)
    This.salesOrganizationList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ActiveUserIndex() As Variant
    Dim i As Long
    For i = 1 To UBound(Me.usersTable, 1)
        If Me.usersTable(i, UsersTableFields.COL_userName) = Me.ActiveUserName Then
            ActiveUserIndex = Me.usersTable(i, UsersTableFields.COL_INDEX)
            Exit Property
        End If
    Next i
End Property

'-------------------------------------------------------------------------

Public Property Get CurrentUpdateIndex() As Long
    CurrentUpdateIndex = This.CurrentUpdateIndex
End Property

Public Property Let CurrentUpdateIndex(ByVal vNewValue As Long)
    This.CurrentUpdateIndex = vNewValue
End Property

'-------------------------------------------------------------------------
'Validations
'-------------------------------------------------------------------------

Public Function IsRecordIDAvailableToUpdate(ByVal TABLE As TablesOfThisApplication, ByVal Index As Long, ByVal RecordID As String) As Variant
    Dim TRow As Long
    Dim ID As String
    Select Case TABLE
        Case TablesOfThisApplication.TABLE_MAINRECORDS
            TRow = AppMethods.GetTargetRowIndex(Me.mainTable, Index, MainTableFields.COL_MAIN_INDEX)
            ID = Me.mainTable(TRow, 2)
        Case TablesOfThisApplication.TABLE_USERS
            TRow = AppMethods.GetTargetRowIndex(Me.usersTable, Index, UsersTableFields.COL_INDEX)
            ID = Me.usersTable(TRow, 2)
    End Select
    If ID = RecordID Then
        IsRecordIDAvailableToUpdate = True
    Else
        IsRecordIDAvailableToUpdate = "Something went wrong! Please try again!"
    End If
End Function

'-------------------------------------------------------------------------
'Public Methods and Functions
'-------------------------------------------------------------------------

'Following function will get new ID for the Table's new record which is _
incrementing and unique as well

Public Function GetNewID(ByVal TABLE As TablesOfThisApplication) As String
    Select Case TABLE
        Case TablesOfThisApplication.TABLE_MAINRECORDS
            If UBound(Me.mainTable, 1) = 1 Then
                GetNewID = INDEX_RECORDID_FIRST
            Else
                GetNewID = VBA.CLng(AppMethods.GetMax(Me.mainTable, MainTableFields.COL_MAIN_recordID)) + 1
            End If
        Case TablesOfThisApplication.TABLE_USERS
            If UBound(Me.usersTable, 1) = 1 Then
                GetNewID = INDEX_USERID_FIRST
            Else
                GetNewID = VBA.CLng(AppMethods.GetMax(Me.usersTable, UsersTableFields.COL_userId)) + 1
            End If
    End Select
End Function

'Purpose:       Following function will pick unique list from the list!
Public Function GetUniqueListFromColumn(ByVal ContextTable As Variant, ByVal TargetColumn As Long) As Variant
    
    'Variable Declarations
    Dim Records             As Variant
    Dim Record              As Variant
    Dim tmp                 As Variant
    Dim i                   As Long
    Dim uName               As String
    Dim TotalMatchingRows   As Long
    
    'Get Generic List of Unique items from the Table's Column
    Records = AppMethods.GetUniqueListFromTableColumn(ContextTable, TargetColumn)
    
'Next step is to filter out uniqueList items based on App's Requirements
    
    'Redimentioning the tmp array
    TotalMatchingRows = UBound(Records)
    ReDim tmp(TotalMatchingRows)
    i = 1
    
    'loop through available list and filter out based on app's criteria
    For Each Record In Records
        If Record <> Empty Then
            If TargetColumn = MainTableFields.COL_MAIN_userID Then
                uName = AppMethods.GetLookupValue(Me.usersTable, Record, UsersTableFields.COL_userId, UsersTableFields.COL_userName)
                If uName <> vbNullString Then
                    tmp(i) = uName
                Else
                    ReDim Preserve tmp(TotalMatchingRows - 1)
                End If
            ElseIf VBA.IsDate(Record) Then
                tmp(i) = VBA.Format$(Record, GetDateFormat)
            Else
                If VBA.IsNumeric(Record) Then
                    tmp(i) = VBA.CDbl(Record)
                Else
                    tmp(i) = Record
                End If
            End If
            i = i + 1
        End If
    Next Record
    
    'return value
    GetUniqueListFromColumn = tmp
    
End Function
