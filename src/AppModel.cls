VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------

Private Type TAppModelComponents
  'Validation Services
    Validator As ValidationServices
  'Data Sources
    mainTable As Variant
    usersTable As Variant
    CurrenciesList As Variant
    UnitOfMeasuresList As Variant
    userTypeList As Variant
    userStatusList As Variant
  'Active User Info
    ActiveUserID As String
    ActiveUserStatus As String
    ActiveUserType As String
    ActiveUserName As String
    ActiveUserPassword As String
End Type

Private this As TAppModelComponents

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

Public Property Get UnitOfMeasuresList() As Variant
    UnitOfMeasuresList = this.UnitOfMeasuresList
End Property

Public Property Let UnitOfMeasuresList(ByVal vNewValue As Variant)
    this.UnitOfMeasuresList = vNewValue
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

Public Property Get ActiveUserIndex() As Variant
    Dim I As Long
    For I = 1 To UBound(Me.usersTable, 1) Step 1
        If Me.usersTable(I, UsersTableFields.COL_userName) = Me.ActiveUserName Then
            ActiveUserIndex = Me.usersTable(I, UsersTableFields.COL_INDEX)
            Exit Property
        End If
    Next I
End Property

'-------------------------------------------------------------------------
'InIt This Model
'-------------------------------------------------------------------------

Public Sub InItModel(ByVal MainTableContent As Variant, _
                    ByVal UsersTableContent As Variant, _
                    ByVal ListOfCurrencies As Variant, _
                    ByVal ListOfUnits As Variant, _
                    ByVal ListofUserTypes As Variant, _
                    ByVal ListofUserStatus As Variant)
    Set Validator = New ValidationServices
    'Tables of The Applications
    mainTable = MainTableContent
    usersTable = UsersTableContent
    'Single Column Lists for DropDown/ComboBoxes
    CurrenciesList = ListOfCurrencies
    UnitOfMeasuresList = ListOfUnits
    userTypeList = ListofUserTypes
    userStatusList = ListofUserStatus
End Sub

'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Validator = Nothing
End Sub

'-------------------------------------------------------------------------
