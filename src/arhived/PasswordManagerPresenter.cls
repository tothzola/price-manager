VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PasswordManagerPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVP.Presenter.PasswordManager")
Option Explicit

'-------------------------------------------------------------------------

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MainView As PriceApprovalView
Attribute MainView.VB_VarHelpID = -1

'-------------------------------------------------------------------------

Public Event CancelPasswordManagerForm()
Public Event OpenPasswordManagerForm(ByVal PMModel As PasswordManagerModel)
Public Event ChangePassword(ByVal NewPassword As String)
Public Event OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)

'-------------------------------------------------------------------------

Private Type TPasswordManagerComponents
    Model As PasswordManagerModel
End Type

Private this As TPasswordManagerComponents

'-------------------------------------------------------------------------

Private Property Get Model() As PasswordManagerModel
    Set Model = this.Model
End Property

Private Property Set Model(ByVal vNewValue As PasswordManagerModel)
    Set this.Model = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt This Presenter with necessary parameters / Constructure procedure
'-------------------------------------------------------------------------

Public Sub InItPasswordManagerForm(ByVal View As PriceApprovalView)
    If Model Is Nothing Then Set Model = New PasswordManagerModel
    If MainView Is Nothing Then Set MainView = View
    RaiseEvent OpenPasswordManagerForm(Model) 'Giving Responcibility of opening frame to a AppPresenter! Only AppPresenter could handle MainView!
End Sub

'Following event handler will help to change the password but
'before it changes the password, it first check if everything is okay?
'if not then it will raise and event that will show the warning message

Private Sub MainView_ChangePassword()
    Dim response As Variant
    response = Model.IsEverythingOkayToChangePassword
    If response = True Then
        RaiseEvent ChangePassword(Model.NewPassword)
    Else
        RaiseEvent OnSomethingWentWrong(response, TYPE_CRITICAL)
    End If
End Sub

'bye-bye procedure

Private Sub MainView_ClosePasswordManagerFrame()
    RaiseEvent CancelPasswordManagerForm
End Sub

'-------------------------------------------------------------------------
'Class Event
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Model = Nothing
    Set MainView = Nothing
End Sub
