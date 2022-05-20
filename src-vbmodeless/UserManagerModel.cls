VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserManagerModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "UserManager"
Option Explicit

'-------------------------------------------------------------------------

Private Type TUserManagerFormModelComponents
    'Validation Objects
    Validator As ValidationServices
    CheckFields(UsersTableFields.COL_userStatus To UsersTableFields.COL_password) As Variant
    'Table
    usersTable As Variant
    'single column lists
    userStatusList As Variant
    userTypesList As Variant
    'Fields
    userIndex As Long
    userID As String
    userStatus As String
    UserType As String
    UserName As String
    userPassword As String
End Type

Private this As TUserManagerFormModelComponents

'-------------------------------------------------------------------------

Private Property Get Validator() As ValidationServices
    Set Validator = this.Validator
End Property

Private Property Set Validator(ByVal vNewValue As ValidationServices)
    Set this.Validator = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get usersTable() As Variant
    usersTable = this.usersTable
End Property

Public Property Let usersTable(ByVal vNewValue As Variant)
    this.usersTable = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userStatusList() As Variant
    userStatusList = this.userStatusList
End Property

Public Property Let userStatusList(ByVal vNewValue As Variant)
    this.userStatusList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userTypesList() As Variant
    userTypesList = this.userTypesList
End Property

Public Property Let userTypesList(ByVal vNewValue As Variant)
    this.userTypesList = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userIndex() As Long
    userIndex = this.userIndex
End Property

Public Property Let userIndex(ByVal vNewValue As Long)
    this.userIndex = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userID() As String
    userID = this.userID
End Property

Public Property Let userID(ByVal vNewValue As String)
    this.userID = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userStatus() As String
    userStatus = this.userStatus
End Property

Public Property Let userStatus(ByVal vNewValue As String)
    this.userStatus = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get UserType() As String
    UserType = this.UserType
End Property

Public Property Let UserType(ByVal vNewValue As String)
    this.UserType = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get UserName() As String
    UserName = this.UserName
End Property

Public Property Let UserName(ByVal vNewValue As String)
    this.UserName = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userPassword() As String
    userPassword = this.userPassword
End Property

Public Property Let userPassword(ByVal vNewValue As String)
    this.userPassword = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt this model
'-------------------------------------------------------------------------

Public Sub InItModel(ByVal usersData As Variant, ByVal StatusOfUser As Variant, ByVal TypesOfUser As Variant)
    If Validator Is Nothing Then Set Validator = New ValidationServices
    usersTable = usersData
    userStatusList = StatusOfUser
    userTypesList = TypesOfUser
    'reset check field array
    'Changed this to approach because earlier code wasn't resetting the each items of the array to false!
    Call Validator.ResetSingleDimensionalArray(this.CheckFields)
End Sub

'-------------------------------------------------------------------------

Public Sub SetPropertiesToNewUserState()
    'Flush All the properties before we start operation
    Call SetPropertiesToNullState
    With Me
        .userIndex = 0
        .userStatus = USERSTATUS_ACTIVE
        .UserType = USERTYPE_CLIENT
    End With
End Sub

Public Sub SetPropertiesToUpdateUserState()
    'Flush All the properties before we start operation
    Call SetPropertiesToNullState
    'update operation begins
    Dim TargetRow As Long
    TargetRow = GetTargetRowIndex(Me.usersTable, Me.userIndex, UsersTableFields.COL_INDEX)
    With Me
        .userID = Me.usersTable(TargetRow, UsersTableFields.COL_userID)
        .UserName = Me.usersTable(TargetRow, UsersTableFields.COL_userName)
        .userPassword = Me.usersTable(TargetRow, UsersTableFields.COL_password)
        .userStatus = Me.usersTable(TargetRow, UsersTableFields.COL_userStatus)
        .UserType = Me.usersTable(TargetRow, UsersTableFields.COL_userType)
    End With
End Sub

Private Sub SetPropertiesToNullState()
    With Me
        .userID = vbNullString
        .UserName = vbNullString
        .userPassword = vbNullString
        .userStatus = vbNullString
        .UserType = vbNullString
    End With
End Sub

'-------------------------------------------------------------------------
'Validation Methods
'-------------------------------------------------------------------------

Public Function IsValidField(ByVal FieldIndex As UsersTableFields) As Boolean
    'Reset CheckField's State to False
    this.CheckFields(FieldIndex) = False
    'Check the validation of field based on field index
    Select Case FieldIndex
        Case UsersTableFields.COL_userStatus
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.userStatus, VBA.TypeName(Me.userStatus))
            If Not IsValidField = True Then Exit Function
       Case UsersTableFields.COL_userType
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.UserType, VBA.TypeName(Me.UserType))
            If Not IsValidField = True Then Exit Function
        Case UsersTableFields.COL_userName
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.UserName, VBA.TypeName(Me.UserName))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = Validator.IsValidStringLength(Me.UserName, 6)
            If Not IsValidField = True Then Exit Function
            'Check 3
            IsValidField = Validator.IsUnique(Me.usersTable, Me.UserName, UsersTableFields.COL_userName)
            If Not IsValidField = True Then Exit Function
        Case UsersTableFields.COL_password
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.userPassword, VBA.TypeName(Me.userPassword))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = Validator.IsValidPasswordString(Me.userPassword)
            If Not IsValidField = True Then Exit Function
    End Select
    'If All Goes Well!
    this.CheckFields(FieldIndex) = True
End Function

Public Function IsEverythingOkayToSaveUser() As Variant
    IsEverythingOkayToSaveUser = Validator.IsEverythingOkay(this.CheckFields, _
                                                            modDataSources.arrListOfColumns_USERS_TABLE, _
                                                            UsersTableFields.COL_userStatus, _
                                                            UsersTableFields.COL_password)
End Function

Private Function GetTargetRowIndex(ByVal TargetTable As Variant, ByVal TargetIndex As Long, ByVal TargetColumn As Long) As Long
    Dim I As Long
    For I = LBound(TargetTable, 1) To UBound(TargetTable, 1)
        If TargetTable(I, TargetColumn) = TargetIndex Then
            GetTargetRowIndex = I
            Exit Function
        End If
    Next I
End Function

'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Validator = Nothing
End Sub