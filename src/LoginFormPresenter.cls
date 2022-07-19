VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoginFormPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PriceApproval.Presenter.LoginForm")
Option Explicit

'-------------------------------------------------------------------------

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MainView As PriceApprovalView
Attribute MainView.VB_VarHelpID = -1

'-------------------------------------------------------------------------

Public Event OpenLoginForm(ByVal LModel As LoginFormModel)
Public Event CancelLoginForm()
Public Event Login()
Public Event ResetPasswordForUser(ByVal targetUserName As String, ByVal TargetEmailAddress As String)
Public Event OnSomethingWentWrong(ByVal Message As String, ByVal typeOfMessage As messageType)

'-------------------------------------------------------------------------

Private Type TLoginFormComponents
    Model As LoginFormModel
End Type

'-------------------------------------------------------------------------

Private this As TLoginFormComponents

'-------------------------------------------------------------------------

Private Property Get Model() As LoginFormModel
    Set Model = this.Model
End Property

Private Property Set Model(ByVal vNewValue As LoginFormModel)
    Set this.Model = vNewValue
End Property

'-------------------------------------------------------------------------

Public Sub InItLoginForm(ByVal View As PriceApprovalView)
    If Model Is Nothing Then Set Model = New LoginFormModel
    If MainView Is Nothing Then Set MainView = View
    RaiseEvent OpenLoginForm(Model) 'Giving Responcibility of opening frame to a AppPresenter! Only AppPresenter could handle MainView!
End Sub

Private Sub MainView_CloseLoginFrame()
    RaiseEvent CancelLoginForm
End Sub

Private Sub MainView_Login()
    RaiseEvent Login
End Sub

Private Sub MainView_ResetPassword(ByVal targetUserName As String, ByVal TargetEmailAddress As String)
    Dim response As Variant
    response = Model.IsValidDataForResettingPassword(targetUserName, TargetEmailAddress)
    If response = True Then
        RaiseEvent ResetPasswordForUser(targetUserName, TargetEmailAddress)
    Else
        RaiseEvent OnSomethingWentWrong(response, TYPE_CRITICAL)
    End If
End Sub

'-------------------------------------------------------------------------
'Class Event
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Model = Nothing
    Set MainView = Nothing
End Sub
