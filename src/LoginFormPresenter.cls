VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoginFormPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "LoginForm"
Option Explicit

'-------------------------------------------------------------------------

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MainView As PAM
Attribute MainView.VB_VarHelpID = -1

'-------------------------------------------------------------------------

Public Event OpenLoginForm(ByVal LModel As LoginFormModel)
Public Event CancelLoginForm()
Public Event Login()

'-------------------------------------------------------------------------

Private Type TLoginFormComponents
    model As LoginFormModel
End Type

'-------------------------------------------------------------------------

Private this As TLoginFormComponents

'-------------------------------------------------------------------------

Private Property Get model() As LoginFormModel
    Set model = this.model
End Property

Private Property Set model(ByVal vNewValue As LoginFormModel)
    Set this.model = vNewValue
End Property

'-------------------------------------------------------------------------

Public Sub InItLoginForm(ByVal View As PAM)
    If model Is Nothing Then Set model = New LoginFormModel
    If MainView Is Nothing Then Set MainView = View
    RaiseEvent OpenLoginForm(model) 'Giving Responcibility of opening frame to a AppPresenter! Only AppPresenter could handle MainView!
End Sub

Private Sub MainView_CloseLoginFrame()
    RaiseEvent CancelLoginForm
End Sub

Private Sub MainView_Login()
    RaiseEvent Login
End Sub

'-------------------------------------------------------------------------
'Class Event
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set model = Nothing
    Set MainView = Nothing
End Sub
