VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PasswordManagerModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "UserManager"
Option Explicit

'-------------------------------------------------------------------------

Private Type TPasswordManagerFormModelComponents
    Validator As ValidationServices
    password As String
    insertedPassword As String
    newPassword As String
    confirmNewPassword As String
End Type

Private this As TPasswordManagerFormModelComponents

'-------------------------------------------------------------------------
'Properties
'-------------------------------------------------------------------------

Public Property Get confirmNewPassword() As String
    confirmNewPassword = this.confirmNewPassword
End Property

Public Property Let confirmNewPassword(ByVal vNewValue As String)
    this.confirmNewPassword = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get newPassword() As String
    newPassword = this.newPassword
End Property

Public Property Let newPassword(ByVal vNewValue As String)
    this.newPassword = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get insertedPassword() As String
    insertedPassword = this.insertedPassword
End Property

Public Property Let insertedPassword(ByVal vNewValue As String)
    this.insertedPassword = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get Validator() As ValidationServices
    Set Validator = this.Validator
End Property

Private Property Set Validator(ByVal vNewValue As ValidationServices)
    Set this.Validator = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get password() As String
    password = this.password
End Property

Public Property Let password(ByVal vNewValue As String)
    this.password = vNewValue
End Property

'-------------------------------------------------------------------------
'Init Method
'-------------------------------------------------------------------------

Public Sub InItModel(ByVal ActiveUserPassword As String)
    If Validator Is Nothing Then Set Validator = New ValidationServices
    Me.password = ActiveUserPassword
End Sub

'Methods and Function

Public Function IsValidField(ByVal FieldIndex As Long) As Boolean
    Select Case FieldIndex
        'Current Password Field
        Case 1
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.insertedPassword, VBA.TypeName(Me.insertedPassword))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = Validator.ComapreTwoStrings(Me.insertedPassword, Me.password)
        'New Password Field
        Case 2
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.newPassword, VBA.TypeName(Me.newPassword))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = Validator.IsValidPasswordString(Me.newPassword)
        'Confirm new Password Field
        Case 3
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.confirmNewPassword, VBA.TypeName(Me.confirmNewPassword))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = Validator.ComapreTwoStrings(Me.newPassword, Me.confirmNewPassword)
    End Select
End Function

Public Function IsEverythingOkayToChangePassword() As Variant
    Dim checkField1 As Boolean
    Dim checkField2 As Boolean
    Dim checkField3 As Boolean
    Dim tmp As String
    checkField1 = IsValidField(1)
    checkField2 = IsValidField(2)
    checkField3 = IsValidField(3)
    If checkField1 = True And checkField2 = True And checkField3 = True Then
        IsEverythingOkayToChangePassword = True
    Else
        tmp = vbNullString
        If checkField1 = False Then
            tmp = tmp & "Current Password" & vbNewLine
        End If
        If checkField2 = False Then
            tmp = tmp & "New Password" & vbNewLine
        End If
        If checkField3 = False Then
            tmp = tmp & "Confirm New Password" & vbNewLine
        End If
        IsEverythingOkayToChangePassword = "Kindly fill the following fields(s) properly to change your Password!" & vbNewLine & vbNewLine & tmp
    End If
End Function

'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Validator = Nothing
End Sub
