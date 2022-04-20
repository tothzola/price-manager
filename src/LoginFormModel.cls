VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoginFormModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------

Private Type TLoginFormModelComponents
    Validator As ValidationServices
    usersData As Variant
    userName As String
    password As String
    TargetIndex As Long
End Type

Private this As TLoginFormModelComponents

'-------------------------------------------------------------------------
'Properties
'-------------------------------------------------------------------------

Private Property Get Validator() As ValidationServices
    Set Validator = this.Validator
End Property

Private Property Set Validator(ByVal vNewValue As ValidationServices)
    Set this.Validator = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get usersData() As Variant
    usersData = this.usersData
End Property

Public Property Let usersData(ByVal vNewValue As Variant)
    this.usersData = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userName() As String
    userName = this.userName
End Property

Public Property Let userName(ByVal vNewValue As String)
    this.userName = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get password() As String
    password = this.password
End Property

Public Property Let password(ByVal vNewValue As String)
    this.password = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get TargetIndex() As Long
    TargetIndex = this.TargetIndex
End Property

Public Property Let TargetIndex(ByVal vNewValue As Long)
    this.TargetIndex = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get userType() As String
    userType = Me.usersData(Me.TargetIndex, UsersTableFields.COL_userType)
End Property

'-------------------------------------------------------------------------

Public Property Get userStatus() As String
    Dim I As Long
    For I = 1 To UBound(Me.usersData, 1) Step 1
        If VBA.UCase$(Me.usersData(I, UsersTableFields.COL_userName)) = VBA.UCase$(Me.userName) Then
            Me.TargetIndex = I
            userStatus = Me.usersData(I, UsersTableFields.COL_userStatus)
            Exit Property
        End If
    Next I
End Property

'-------------------------------------------------------------------------

Public Property Get userID() As String
    userID = Me.usersData(TargetIndex, UsersTableFields.COL_userID)
End Property

'-------------------------------------------------------------------------
'InIt Method
'-------------------------------------------------------------------------

Public Sub InItModel(ByVal UsersTableContent As Variant)
    If Validator Is Nothing Then Set Validator = New ValidationServices
    Me.usersData = UsersTableContent
End Sub

'-------------------------------------------------------------------------
'Methods & Validation
'-------------------------------------------------------------------------

Public Function IsValidUsername() As Boolean
    IsValidUsername = Validator.IsNotNullValue(Me.userName, VBA.TypeName(Me.userName))
End Function

Public Function IsValidPassword() As Boolean
    IsValidPassword = Validator.IsNotNullValue(Me.password, VBA.TypeName(Me.password))
End Function

Public Function IsUserAuthorized() As Variant
    IsUserAuthorized = Validator.IsUserAuthorized(Me.usersData, Me.userName, Me.password, UsersTableFields.COL_userName, UsersTableFields.COL_password)
End Function

'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Validator = Nothing
End Sub