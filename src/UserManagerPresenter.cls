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
Public Event AddNewUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal UMModel As UserManagerModel)
Public Event UpdateUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal UMModel As UserManagerModel)
Public Event DeleteUserFromUserManagerForm(ByVal TypeOfOperation As CRUDOperations, ByVal UMModel As UserManagerModel)
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

Public Sub InItUserManagerForm(ByVal View As PAM)
    If model Is Nothing Then Set model = New UserManagerModel
    If MainView Is Nothing Then Set MainView = View
    RaiseEvent OpenUserManagerForm(model)
End Sub

Private Sub MainView_CloseUserManagerFrame()
    RaiseEvent CloseUserManagerForm
End Sub

Private Sub MainView_DoCRUDOperationForUserManager(ByVal TypeOfOperation As CRUDOperations)
    Dim response As Variant
    With model
        'Validation
        Select Case TypeOfOperation
        Case CRUDOperations.CRUD_OPERATION_ADDNEW
            response = .IsEverythingOkayToSaveUser(OPERATION_NEW)
        Case CRUDOperations.CRUD_OPERATION_UPDATE
            response = .IsEverythingOkayToSaveUser(OPERATION_UPDATE)
        Case CRUDOperations.CRUD_OPERATION_DELETE
            response = .IsEverythingOkayToSaveUser(OPERATION_DELETE)
        End Select
        'Now based on the validation go for the operation
        If response = True Then
            Select Case TypeOfOperation
            Case CRUDOperations.CRUD_OPERATION_ADDNEW
                RaiseEvent AddNewUserFromUserManagerForm(TypeOfOperation, model)
            Case CRUDOperations.CRUD_OPERATION_UPDATE
                RaiseEvent UpdateUserFromUserManagerForm(TypeOfOperation, model)
            Case CRUDOperations.CRUD_OPERATION_DELETE
                RaiseEvent DeleteUserFromUserManagerForm(TypeOfOperation, model)
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

'-------------------------------------------------------------------------
'Class Event
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set model = Nothing
    Set MainView = Nothing
End Sub
