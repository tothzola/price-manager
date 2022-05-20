VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PasswordManagerModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PasswordManager"
Option Explicit

'-------------------------------------------------------------------------

Private Type TPasswordManagerFormModelComponents
    'Validation Objects
    Validator As ValidationServices
    CheckFields(PasswordManagerFields.PM_CURRENT_PASSWORD To PasswordManagerFields.PM_CONFIRM_NEW_PASSWORD) As Variant
    'field Properties
    Password As String
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

Public Property Get Password() As String
    Password = this.Password
End Property

Public Property Let Password(ByVal vNewValue As String)
    this.Password = vNewValue
End Property

'-------------------------------------------------------------------------
'Init Method
'-------------------------------------------------------------------------

Public Sub InItModel(ByVal ActiveUserPassword As String)
    If Validator Is Nothing Then Set Validator = New ValidationServices
    Me.Password = ActiveUserPassword
    'reset check field array
    'Changed this to approach because earlier code wasn't resetting the each items of the array to false!
    Call Validator.ResetSingleDimensionalArray(this.CheckFields)
End Sub

'-------------------------------------------------------------------------
'Validation Method
'-------------------------------------------------------------------------

Public Function IsEverythingOkayToChangePassword() As Variant
    IsEverythingOkayToChangePassword = Validator.IsEverythingOkay(this.CheckFields, _
                                                                modDataSources.arrListOfFields_PASSWORD_MANAGER, _
                                                                PasswordManagerFields.PM_CURRENT_PASSWORD, _
                                                                PasswordManagerFields.PM_CONFIRM_NEW_PASSWORD)
End Function

Public Function IsValidField(ByVal FieldIndex As PasswordManagerFields) As Boolean
    'Reset CheckField's State to False
    this.CheckFields(FieldIndex) = False
    'Check the validation of field based on field index
    Select Case FieldIndex
        Case PasswordManagerFields.PM_CURRENT_PASSWORD
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.insertedPassword, VBA.TypeName(Me.insertedPassword))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = Validator.ComapreTwoStrings(Me.insertedPassword, Me.Password)
            If Not IsValidField = True Then Exit Function
        Case PasswordManagerFields.PM_NEW_PASSWORD
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.newPassword, VBA.TypeName(Me.newPassword))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = Validator.IsValidPasswordString(Me.newPassword)
            If Not IsValidField = True Then Exit Function
        Case PasswordManagerFields.PM_CONFIRM_NEW_PASSWORD
            'Check 1
            IsValidField = Validator.IsNotNullValue(Me.confirmNewPassword, VBA.TypeName(Me.confirmNewPassword))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = Validator.ComapreTwoStrings(Me.newPassword, Me.confirmNewPassword)
            If Not IsValidField = True Then Exit Function
    End Select
    'If All Goes Well!
    this.CheckFields(FieldIndex) = True
End Function

'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Validator = Nothing
End Sub
