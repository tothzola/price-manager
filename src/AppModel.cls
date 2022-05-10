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
  'Active User Info
    ActiveUserID As String
    ActiveUserStatus As String
    ActiveUserType As String
    ActiveUserName As String
    ActiveUserPassword As String
  'Current Parameters
    CurrentUpdateIndex As Long
End Type

Private this As TAppModelComponents

'-------------------------------------------------------------------------
'Properties
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

Public Property Get ActiveUserIndex() As Variant
    Dim I As Long
    For I = 1 To UBound(Me.usersTable, 1)
        If Me.usersTable(I, UsersTableFields.COL_userName) = Me.ActiveUserName Then
            ActiveUserIndex = Me.usersTable(I, UsersTableFields.COL_INDEX)
            Exit Property
        End If
    Next I
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
                    ByVal ListOfRecordStatus As Variant)
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
End Sub

'-------------------------------------------------------------------------
'Validations
'-------------------------------------------------------------------------

Public Function IsRecordIDAvailableToUpdate(ByVal TABLE As TablesOfThisApplication, ByVal index As Long, ByVal RecordID As String) As Variant
    Dim TRow As Long
    Dim ID As String
    TRow = index + 1
    Select Case TABLE
        Case TablesOfThisApplication.TABLE_MAINRECORDS
            ID = Me.mainTable(TRow, 2)
        Case TablesOfThisApplication.Table_Users
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
                GetNewID = VBA.CLng(Me.mainTable(UBound(Me.mainTable, 1), 2)) + 1
            End If
        Case TablesOfThisApplication.Table_Users
            If UBound(Me.usersTable, 1) = 1 Then
                GetNewID = INDEX_USERID_FIRST
            Else
                GetNewID = VBA.CLng(Me.usersTable(UBound(Me.usersTable, 1), 2)) + 1
            End If
    End Select
End Function

'Following function will pick unique list from the list!
Public Function GetUniqueListFromColumn(ByVal ContextTable As Variant, ByVal targetColumn As Long) As Variant
    'Driving Paramters
    Dim Records()               As String
    Dim Record                  As Variant
    Dim MatchedRecords          As String
    Dim TotalMatchingRows       As Long
    Dim firstRow                As Long
    Dim lastRow                 As Long
    Dim I                       As Long
    Dim tmp                     As Variant
    Dim dict                    As Dictionary
    'Get Driving parameters
    firstRow = LBound(ContextTable, 1)
    lastRow = UBound(ContextTable, 1)
    ReDim tmp(0)
    'Get matching Rows!
    MatchedRecords = vbNullString
    TotalMatchingRows = 0
    Set dict = New Dictionary
    For I = firstRow + 1 To lastRow
        If Not dict.Exists(ContextTable(I, targetColumn)) Then
            dict.Add ContextTable(I, targetColumn), ContextTable(I, targetColumn)
            MatchedRecords = MatchedRecords & I & SEPERATOR_ITEM
            TotalMatchingRows = TotalMatchingRows + 1
        End If
    Next I
    'remove last seperator from the matchedrecords variable
    Set dict = Nothing
    If MatchedRecords = vbNullString Then
        GetUniqueListFromColumn = tmp
        Exit Function
    Else
        MatchedRecords = VBA.Left$(MatchedRecords, VBA.Len(MatchedRecords) - VBA.Len(SEPERATOR_ITEM))
    End If
    'Redimentioning the tmp array
    ReDim tmp(TotalMatchingRows)
    'Transferring Data
    Records = VBA.Split(MatchedRecords, SEPERATOR_ITEM)
    I = 1
    For Each Record In Records
        tmp(I) = ContextTable(Record, targetColumn)
        I = I + 1
    Next Record
    'return value
    GetUniqueListFromColumn = SortOneDimensionalArrayWithNumbers(tmp)
End Function

'Following function will Sort One Dimensional Array in Ascending Order !
Private Function SortOneDimensionalArrayWithNumbers(ByVal arr As Variant) As Variant
    Dim FirstItemIndex  As Long
    Dim LastItemIndex   As Long
    Dim tmp             As Long
    Dim I               As Long
    Dim J               As Long
    FirstItemIndex = LBound(arr)
    LastItemIndex = UBound(arr)
    For I = FirstItemIndex To LastItemIndex - 1
        For J = I + 1 To LastItemIndex
            If Val(arr(I)) > Val(arr(J)) Then
                tmp = arr(J)
                arr(J) = arr(I)
                arr(I) = tmp
            End If
        Next J
    Next I
    SortOneDimensionalArrayWithNumbers = arr
End Function

'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Validator = Nothing
End Sub


