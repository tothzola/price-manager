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
    Validator As ValidationServices
    'Table
    usersTable As Variant
    'single column lists
    userStatusList As Variant
    userTypesList As Variant
    'Fields
    userIndex As Long
    userID As String
    userStatus As String
    userType As String
    userName As String
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

Public Property Get userType() As String
    userType = this.userType
End Property

Public Property Let userType(ByVal vNewValue As String)
    this.userType = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userName() As String
    userName = this.userName
End Property

Public Property Let userName(ByVal vNewValue As String)
    this.userName = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userPassword() As String
    userPassword = this.userPassword
End Property

Public Property Let userPassword(ByVal vNewValue As String)
    this.userPassword = vNewValue
End Property

'-------------------------------------------------------------------------


'-------------------------------------------------------------------------
'InIt this model
'-------------------------------------------------------------------------

Public Sub InItModel(ByVal usersData As Variant, ByVal StatusOfUser As Variant, ByVal TypesOfUser As Variant)
    If Validator Is Nothing Then Set Validator = New ValidationServices
    usersTable = usersData
    userStatusList = StatusOfUser
    userTypesList = TypesOfUser
End Sub

'-------------------------------------------------------------------------

Public Sub SetPropertiesToNewUserState()
    'Flush All the properties before we start operation
    Call SetPropertiesToNullState
    With Me
        .userIndex = 0
        If UBound(Me.usersTable, 1) = 1 Then
            .userID = INDEX_USERID_FIRST
        Else
            .userID = VBA.CLng(Me.usersTable(UBound(Me.usersTable, 1), UsersTableFields.COL_userID)) + 1
        End If
        .userStatus = USERSTATUS_ACTIVE
        .userType = USERTYPE_CLIENT
    End With
End Sub

Public Sub SetPropertiesToUpdateUserState()
    'Flush All the properties before we start operation
    Call SetPropertiesToNullState
    'update operation begins
    Dim TargetRow As Long
    TargetRow = Me.userIndex
    With Me
        .userID = Me.usersTable(TargetRow, UsersTableFields.COL_userID)
        .userName = Me.usersTable(TargetRow, UsersTableFields.COL_userName)
        .userPassword = Me.usersTable(TargetRow, UsersTableFields.COL_password)
        .userStatus = Me.usersTable(TargetRow, UsersTableFields.COL_userStatus)
        .userType = Me.usersTable(TargetRow, UsersTableFields.COL_userType)
    End With
End Sub

Private Sub SetPropertiesToNullState()
    With Me
        .userID = vbNullString
        .userName = vbNullString
        .userPassword = vbNullString
        .userStatus = vbNullString
        .userType = vbNullString
    End With
End Sub

'-------------------------------------------------------------------------
'Validation Methods
'-------------------------------------------------------------------------

'3. userStatus  4. userType    5. userName    6. password
Public Function IsValidField(ByVal FieldIndex As UsersTableFields) As Boolean
    Select Case FieldIndex
    'USER STATUS
        Case 3
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.userStatus, VBA.TypeName(Me.userStatus))
    'USER TYPE
        Case 4
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.userType, VBA.TypeName(Me.userType))
    'USERNAME
        Case 5
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.userName, VBA.TypeName(Me.userName))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = Validator.IsValidStringLength(Me.userName, 6)
    'PASSWORD
        Case 6
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.userPassword, VBA.TypeName(Me.userPassword))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = Validator.IsValidPasswordString(Me.userPassword)
    End Select
End Function

Public Function IsEverythingOkayToSaveUser() As Variant
    Dim checkField1 As Boolean
    Dim checkField2 As Boolean
    Dim checkField3 As Boolean
    Dim checkField4 As Boolean
    Dim tmp As String
    checkField1 = IsValidField(COL_userStatus)
    checkField2 = IsValidField(COL_userType)
    checkField3 = IsValidField(COL_userName)
    checkField4 = IsValidField(COL_password)
    If checkField1 = True And checkField2 = True And checkField3 = True And checkField4 = True Then
        IsEverythingOkayToSaveUser = True
    Else
        tmp = vbNullString
        If checkField1 = False Then
            tmp = tmp & "User Status" & vbNewLine
        End If
        If checkField2 = False Then
            tmp = tmp & "User Type" & vbNewLine
        End If
        If checkField3 = False Then
            tmp = tmp & "Username" & vbNewLine
        End If
        If checkField4 = False Then
            tmp = tmp & "Password" & vbNewLine
        End If
        IsEverythingOkayToSaveUser = "Kindly fill the following fields(s) to proceed further!" & vbNewLine & vbNewLine & tmp
    End If
End Function


'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Validator = Nothing
End Sub
