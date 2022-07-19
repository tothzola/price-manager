VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoginFormModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PriceApproval.Model.LoginForm")
Option Explicit

'-------------------------------------------------------------------------

Private Type TLoginFormModelComponents
    'Table
    usersData As Variant
    'Fields
    UserName As String
    Password As String
    'Attributes
    TargetIndex As Long
    
End Type

Private this As TLoginFormModelComponents

'-------------------------------------------------------------------------

Public Property Get usersData() As Variant
    usersData = this.usersData
End Property

Public Property Let usersData(ByVal vNewValue As Variant)
    this.usersData = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get UserName() As String
    UserName = this.UserName
End Property

Public Property Let UserName(ByVal vNewValue As String)
    this.UserName = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get Password() As String
    Password = this.Password
End Property

Public Property Let Password(ByVal vNewValue As String)
    this.Password = vNewValue
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
    Dim i As Long
    For i = 1 To UBound(Me.usersData, 1)
        If UCase$(Me.usersData(i, UsersTableFields.COL_userName)) = UCase$(Me.UserName) Then
            Me.UserName = Me.usersData(i, UsersTableFields.COL_userName)
            Me.TargetIndex = i
            userStatus = Me.usersData(i, UsersTableFields.COL_userStatus)
            Exit Property
        End If
    Next i
End Property

'-------------------------------------------------------------------------

Public Property Get userEmail() As String
    Dim i As Long
    For i = 1 To UBound(Me.usersData, 1)
        If UCase$(Me.usersData(i, UsersTableFields.COL_userName)) = UCase$(Me.UserName) Then
            Me.UserName = Me.usersData(i, UsersTableFields.COL_userName)
            Me.TargetIndex = i
            userEmail = Me.usersData(i, UsersTableFields.COL_email)
            Exit Property
        End If
    Next i
End Property

'-------------------------------------------------------------------------

Public Property Get userID() As String
    userID = Me.usersData(TargetIndex, UsersTableFields.COL_userId)
End Property

'-------------------------------------------------------------------------
'InIt Method
'-------------------------------------------------------------------------

Public Sub InItModel(ByVal UsersTableContent As Variant)
    Me.usersData = UsersTableContent
End Sub

'-------------------------------------------------------------------------
'Methods & Validation
'-------------------------------------------------------------------------

Public Function IsValidUsername() As Boolean
    IsValidUsername = ValidationServices.IsNotNullValue(Me.UserName, VBA.TypeName(Me.UserName))
End Function

Public Function IsValidPassword() As Boolean
    IsValidPassword = ValidationServices.IsNotNullValue(Me.Password, VBA.TypeName(Me.Password))
End Function

Public Function IsUserAuthorized() As Variant
    IsUserAuthorized = ValidationServices.IsUserAuthorized(Me.usersData, Me.UserName, Me.Password, UsersTableFields.COL_userName, UsersTableFields.COL_password)
End Function

Public Function IsValidDataForResettingPassword(ByVal tUserName As String, ByVal tEmail As String) As Variant

    Dim i As Long
    For i = LBound(Me.usersData, 1) + 1 To UBound(Me.usersData, 1)
        If UCase$(Me.usersData(i, UsersTableFields.COL_userName)) = UCase$(tUserName) Then
            If Me.usersData(i, UsersTableFields.COL_userStatus) = USERSTATUS_INACTIVE Then
                IsValidDataForResettingPassword = "User is InActive! Cannot reset password. Kindly Contact Admin for more information"
                Exit Function
            End If
            If UCase$(Me.usersData(i, UsersTableFields.COL_email)) = UCase$(tEmail) Then
                IsValidDataForResettingPassword = True
            Else
                IsValidDataForResettingPassword = "Username and Email doesn't not match. Please try again."
            End If
            Exit Function
        End If
    Next i
    IsValidDataForResettingPassword = "Username is Invalid"
    
End Function

