VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PriceFormPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------

Private WithEvents MainView As PAM
Attribute MainView.VB_VarHelpID = -1

'-------------------------------------------------------------------------

Public Event OpenPriceForm(ByVal PFModel As PriceFormModel)
Public Event ClosePriceForm()
Public Event ResetPriceForm()
Public Event AddNewUserFromPriceForm(ByVal TypeOfOperation As CRUDOperations)
Public Event UpdateUserFromPriceForm(ByVal TypeOfOperation As CRUDOperations)
Public Event DeleteUserFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal index As Long)
Public Event ApproveUserFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal index As Long)
Public Event RejectUserFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal index As Long)
Public Event OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)

'-------------------------------------------------------------------------

Private Type TPriceFormComponents
    model As PriceFormModel
End Type

Private this As TPriceFormComponents

'-------------------------------------------------------------------------

Private Property Get model() As PriceFormModel
    Set model = this.model
End Property

Private Property Set model(ByVal vNewValue As PriceFormModel)
    Set this.model = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt This Presenter with necessary parameters / Constructure procedure
'-------------------------------------------------------------------------

Public Sub InItPriceForm(ByVal view As PAM, ByVal mainData As Variant)
    If model Is Nothing Then Set model = New PriceFormModel
    If MainView Is Nothing Then Set MainView = view
    Call model.InItModel(mainData)
    RaiseEvent OpenPriceForm(model)
End Sub

Private Sub MainView_ClosePriceFormFrame()
    RaiseEvent ClosePriceForm
End Sub


