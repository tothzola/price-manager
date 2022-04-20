VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserManagerPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "UserManager"
Option Explicit

'-------------------------------------------------------------------------

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MainView As PAM
Attribute MainView.VB_VarHelpID = -1

'-------------------------------------------------------------------------

Public Event OpenUserManagerForm(ByVal UMModel As UserManagerModel)
Public Event CloseUserManagerForm()
Public Event ResetUserManagerForm()
Public Event AddNewUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, _
                                            ByVal userID As String, _
                                            ByVal userStatus As String, _
                                            ByVal userType As String, _
                                            ByVal userName As String, _
                                            ByVal userPassword As String)
                                            
Public Event UpdateUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, _
                                            ByVal index As Long, _
                                            ByVal userID As String, _
                                            ByVal userStatus As String, _
                                            ByVal userType As String, _
                                            ByVal userName As String, _
                                            ByVal userPassword As String)
Public Event DeleteUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal index As Long)
Public Event UpdateUserManagerFormRecord()
Public Event OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)

'-------------------------------------------------------------------------

Private Type TUserManagerComponents
    model As UserManagerModel
End Type

Private this As TUserManagerComponents

'-------------------------------------------------------------------------

Private Property Get model() As UserManagerModel
    Set model = this.model
End Property

Private Property Set model(ByVal vNewValue As UserManagerModel)
    Set this.model = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt This Presenter with necessary parameters / Constructure procedure
'-------------------------------------------------------------------------

Public Sub InItUserManagerForm(ByVal view As PAM, ByVal usersData As Variant, ByVal StatusOfUser As Variant, ByVal TypesOfUser As Variant)
    If model Is Nothing Then Set model = New UserManagerModel
    If MainView Is Nothing Then Set MainView = view
    Call model.InItModel(usersData, StatusOfUser, TypesOfUser)
    RaiseEvent OpenUserManagerForm(model)
End Sub

Private Sub MainView_CloseUserManagerFrame()
    RaiseEvent CloseUserManagerForm
End Sub

Private Sub MainView_DoCRUDOperationForUserManager(ByVal TypeOfOperation As CRUDOperations)
    Dim response As Variant
    With model
        'Validation
        response = .IsEverythingOkayToSaveUser
        'Now based on the validation go for the operation
        If response = True Then
            Select Case TypeOfOperation
                Case CRUDOperations.CRUD_OPERATION_ADDNEW
                    RaiseEvent AddNewUserFromUserManagerForm(TypeOfOperation, .userID, .userStatus, .userType, .userName, .userPassword)
                Case CRUDOperations.CRUD_OPERATION_UPDATE
                    RaiseEvent UpdateUserFromUserManagerForm(TypeOfOperation, .userIndex - 1, .userID, .userStatus, .userType, .userName, .userPassword)
                Case CRUDOperations.CRUD_OPERATION_DELETE
                    RaiseEvent DeleteUserFromUserManagerForm(TypeOfOperation, .userIndex - 1)
            End Select
        Else
            RaiseEvent OnSomethingWentWrong(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub MainView_ResetUserManagerFrame()
    RaiseEvent ResetUserManagerForm
End Sub

Private Sub MainView_UpdateUserManagerFrameRecord()
    RaiseEvent UpdateUserManagerFormRecord
End Sub

