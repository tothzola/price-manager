VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "AppObjects"
Option Explicit

'-------------------------------------------------------------------------
'Private Type Componenets / Enums
'-------------------------------------------------------------------------

Private Type TAppModelComponents
  'Validation Services
    Validator As ValidationServices
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

Private this As TAppModelComponents

'-------------------------------------------------------------------------
'Properties
'-------------------------------------------------------------------------

Public Property Get ActiveUserEmail() As String
    ActiveUserEmail = this.ActiveUserEmail
End Property

Public Property Let ActiveUserEmail(ByVal vNewValue As String)
    this.ActiveUserEmail = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ActiveUserPassword() As String
    ActiveUserPassword = this.ActiveUserPassword
End Property

Public Property Let ActiveUserPassword(ByVal vNewValue As String)
    this.ActiveUserPassword = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ActiveUserName() As String
    ActiveUserName = this.ActiveUserName
End Property

Public Property Let ActiveUserName(ByVal vNewValue As String)
    this.ActiveUserName = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ActiveUserType() As String
    ActiveUserType = this.ActiveUserType
End Property

Public Property Let ActiveUserType(ByVal vNewValue As String)
    this.ActiveUserType = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ActiveUserStatus() As String
    ActiveUserStatus = this.ActiveUserStatus
End Property

Public Property Let ActiveUserStatus(ByVal vNewValue As String)
    this.ActiveUserStatus = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get ActiveUserID() As String
    ActiveUserID = this.ActiveUserID
End Property

Public Property Let ActiveUserID(ByVal vNewValue As String)
    this.ActiveUserID = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get Validator() As ValidationServices
    Set Validator = this.Validator
End Property

Private Property Set Validator(ByVal vNewValue As ValidationServices)
    Set this.Validator = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get mainTable() As Variant
    mainTable = this.mainTable
End Property

Public Property Let mainTable(ByVal vNewValue As Variant)
    this.mainTable = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get usersTable() As Variant
    usersTable = this.usersTable
End Property

Public Property Let usersTable(ByVal vNewValue As Variant)
    this.usersTable = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get CurrenciesList() As Variant
    CurrenciesList = this.CurrenciesList
End Property

Public Property Let CurrenciesList(ByVal vNewValue As Variant)
    this.CurrenciesList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get unitOfMeasuresList() As Variant
    unitOfMeasuresList = this.unitOfMeasuresList
End Property

Public Property Let unitOfMeasuresList(ByVal vNewValue As Variant)
    this.unitOfMeasuresList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userTypeList() As Variant
    userTypeList = this.userTypeList
End Property

Public Property Let userTypeList(ByVal vNewValue As Variant)
    this.userTypeList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userStatusList() As Variant
    userStatusList = this.userStatusList
End Property

Public Property Let userStatusList(ByVal vNewValue As Variant)
    this.userStatusList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get recordStatusList() As Variant
    recordStatusList = this.recordStatusList
End Property

Public Property Let recordStatusList(ByVal vNewValue As Variant)
    this.recordStatusList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get distributionChannelsList() As Variant
    distributionChannelsList = this.distributionChannelsList
End Property

Public Property Let distributionChannelsList(ByVal vNewValue As Variant)
    this.distributionChannelsList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get salesOrganizationList() As Variant
    salesOrganizationList = this.salesOrganizationList
End Property

Public Property Let salesOrganizationList(ByVal vNewValue As Variant)
    this.salesOrganizationList = vNewValue
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
    CurrentUpdateIndex = this.CurrentUpdateIndex
End Property

Public Property Let CurrentUpdateIndex(ByVal vNewValue As Long)
    this.CurrentUpdateIndex = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt This Model
'-------------------------------------------------------------------------

Public Sub InItModel(ByVal MainTableContent As Variant, _
                    ByVal UsersTableContent As Variant, _
                    ByVal ListofCurrencies As Variant, _
                    ByVal ListOfUnits As Variant, _
                    ByVal ListofUserTypes As Variant, _
                    ByVal ListofUserStatus As Variant, _
                    ByVal ListOfRecordStatus As Variant, _
                    ByVal ListOfSalesOrganizations As Variant, _
                    ByVal ListOfDistributionChannels As Variant)
    Set Validator = New ValidationServices
    'Tables of The Applications
    Me.mainTable = MainTableContent
    Me.usersTable = UsersTableContent
    'Single Column Lists for DropDown/ComboBoxes
    Me.CurrenciesList = ListofCurrencies
    Me.unitOfMeasuresList = ListOfUnits
    Me.userTypeList = ListofUserTypes
    Me.userStatusList = ListofUserStatus
    Me.recordStatusList = ListOfRecordStatus
    Me.salesOrganizationList = ListOfSalesOrganizations
    Me.distributionChannelsList = ListOfDistributionChannels
End Sub

'-------------------------------------------------------------------------
'Validations
'-------------------------------------------------------------------------

Public Function IsRecordIDAvailableToUpdate(ByVal TABLE As TablesOfThisApplication, ByVal index As Long, ByVal RecordID As String) As Variant
    Dim TRow As Long
    Dim ID As String
    Select Case TABLE
        Case TablesOfThisApplication.TABLE_MAINRECORDS
            TRow = GetMethod.GetTargetRowIndex(Me.mainTable, index, MainTableFields.COL_MAIN_INDEX)
            ID = Me.mainTable(TRow, 2)
        Case TablesOfThisApplication.TABLE_USERS
            TRow = GetMethod.GetTargetRowIndex(Me.usersTable, index, UsersTableFields.COL_INDEX)
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
                GetNewID = VBA.CLng(GetMethod.GetMax(Me.mainTable, MainTableFields.COL_MAIN_recordID)) + 1
            End If
        Case TablesOfThisApplication.TABLE_USERS
            If UBound(Me.usersTable, 1) = 1 Then
                GetNewID = INDEX_USERID_FIRST
            Else
                GetNewID = VBA.CLng(GetMethod.GetMax(Me.usersTable, UsersTableFields.COL_userId)) + 1
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
    Dim uname               As String
    Dim TotalMatchingRows   As Long
    
    'Get Generic List of Unique items from the Table's Column
    Records = GetMethod.GetUniqueListFromTableColumn(ContextTable, TargetColumn)
    
'Next step is to filter out uniqueList items based on App's Requirements
    
    'Redimentioning the tmp array
    TotalMatchingRows = UBound(Records)
    ReDim tmp(TotalMatchingRows)
    i = 1
    
    'loop through available list and filter out based on app's criteria
    For Each Record In Records
        If Record <> Empty Then
            If TargetColumn = MainTableFields.COL_MAIN_userID Then
                uname = GetMethod.GetLookupValue(Me.usersTable, Record, UsersTableFields.COL_userId, UsersTableFields.COL_userName)
                If uname <> vbNullString Then
                    tmp(i) = uname
                Else
                    ReDim Preserve tmp(TotalMatchingRows - 1)
                End If
            ElseIf VBA.IsDate(Record) Then
                'patchDateformat
                'tmp(i) = VBA.format$(GetMethod.GetDateFromUSDate(Record), DATEFORMAT_FRONTEND())
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

'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Validator = Nothing
End Sub


