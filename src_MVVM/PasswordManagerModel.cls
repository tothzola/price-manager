VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PasswordManagerModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.Model.PasswordManager")
Option Explicit

'-------------------------------------------------------------------------

Private Type TPasswordManagerFormModelComponents
    CheckFields(PasswordManagerFields.PM_CURRENT_PASSWORD To PasswordManagerFields.PM_CONFIRM_NEW_PASSWORD) As Variant
    'field Properties
    Password As String
    insertedPassword As String
    NewPassword As String
    confirmNewPassword As String
    
End Type

Private This As TPasswordManagerFormModelComponents

'-------------------------------------------------------------------------
'Properties
'-------------------------------------------------------------------------

Public Property Get confirmNewPassword() As String
    confirmNewPassword = This.confirmNewPassword
End Property

Public Property Let confirmNewPassword(ByVal vNewValue As String)
    This.confirmNewPassword = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get NewPassword() As String
    NewPassword = This.NewPassword
End Property

Public Property Let NewPassword(ByVal vNewValue As String)
    This.NewPassword = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get insertedPassword() As String
    insertedPassword = This.insertedPassword
End Property

Public Property Let insertedPassword(ByVal vNewValue As String)
    This.insertedPassword = vNewValue
End Property

'-------------------------------------------------------------------------

Public Property Get Password() As String
    Password = This.Password
End Property

Public Property Let Password(ByVal vNewValue As String)
    This.Password = vNewValue
End Property

'-------------------------------------------------------------------------
'Init Method
'-------------------------------------------------------------------------

Public Sub InItModel(ByVal ActiveUserPassword As String)
    Me.Password = ActiveUserPassword
    'reset check field array
    'Changed this to approach because earlier code wasn't resetting the each items of the array to false!
    Call AppMethods.ResetSingleDimensionalArray(This.CheckFields)
End Sub

'-------------------------------------------------------------------------
'Validation Method
'-------------------------------------------------------------------------

Public Function IsEverythingOkayToChangePassword() As Variant
    IsEverythingOkayToChangePassword = ValidationServices.IsEverythingOkay(This.CheckFields, _
                                                                DataResources.arrListOfFields_PASSWORD_MANAGER, _
                                                                PasswordManagerFields.PM_CURRENT_PASSWORD, _
                                                                PasswordManagerFields.PM_CONFIRM_NEW_PASSWORD)
End Function

Public Function IsValidField(ByVal FieldIndex As PasswordManagerFields) As Boolean
    'Reset CheckField's State to False
    This.CheckFields(FieldIndex) = False
    'Check the validation of field based on field index
    Select Case FieldIndex
        Case PasswordManagerFields.PM_CURRENT_PASSWORD
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.insertedPassword, VBA.TypeName(Me.insertedPassword))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = ValidationServices.ComapreTwoStrings(Me.insertedPassword, Me.Password)
            If Not IsValidField = True Then Exit Function
            
        Case PasswordManagerFields.PM_NEW_PASSWORD
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.NewPassword, VBA.TypeName(Me.NewPassword))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = ValidationServices.IsValidPasswordString(Me.NewPassword)
            If Not IsValidField = True Then Exit Function
            
        Case PasswordManagerFields.PM_CONFIRM_NEW_PASSWORD
            'Check 1
            IsValidField = ValidationServices.IsNotNullValue(Me.confirmNewPassword, VBA.TypeName(Me.confirmNewPassword))
            If Not IsValidField = True Then Exit Function
            'Check 2
            IsValidField = ValidationServices.ComapreTwoStrings(Me.NewPassword, Me.confirmNewPassword)
            If Not IsValidField = True Then Exit Function
            
    End Select
    'If All Goes Well!
    This.CheckFields(FieldIndex) = True
End Function


