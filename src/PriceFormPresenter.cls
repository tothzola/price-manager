VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PriceFormPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("AppObjects.Presenter.PriceForm")
Option Explicit

'-------------------------------------------------------------------------

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MainView As PriceApprovalView
Attribute MainView.VB_VarHelpID = -1

'-------------------------------------------------------------------------

Public Event OpenPriceForm(ByVal PFModel As PriceFormModel, ByVal operation As FormOperation)
Public Event ClosePriceForm()
Public Event ResetPriceForm()
Public Event AddNewRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
Public Event UpdateRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
Public Event DeleteRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
Public Event RejectRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
Public Event ApproveRecordFromPriceForm(ByVal TypeOfOperation As CRUDOperations, ByVal PFModel As PriceFormModel)
Public Event OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)

'-------------------------------------------------------------------------

Private Type TPriceFormComponents
    Model As PriceFormModel
End Type

Private this As TPriceFormComponents

'-------------------------------------------------------------------------

Private Property Get Model() As PriceFormModel
    Set Model = this.Model
End Property

Private Property Set Model(ByVal vNewValue As PriceFormModel)
    Set this.Model = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt This Presenter with necessary parameters / Constructure procedure
'-------------------------------------------------------------------------

Public Sub InItPriceForm(ByVal View As PriceApprovalView, ByVal operation As FormOperation)
    If Model Is Nothing Then Set Model = New PriceFormModel
    If MainView Is Nothing Then Set MainView = View
    RaiseEvent OpenPriceForm(Model, operation)
End Sub

'-------------------------------------------------------------------------
'Events Procedure
'-------------------------------------------------------------------------

Private Sub MainView_DoCRUDOperationForPriceForm(ByVal TypeOfOperation As CRUDOperations)
    Dim response As Variant
    With Model
        'Validation
        response = .IsEverythingOkayToSaveRecord
        'Now based on the validation go for the operation
        If response = True Then
            Select Case TypeOfOperation
                Case CRUDOperations.CRUD_OPERATION_ADDNEW
                    RaiseEvent AddNewRecordFromPriceForm(TypeOfOperation, Model)
                Case CRUDOperations.CRUD_OPERATION_UPDATE
                    RaiseEvent UpdateRecordFromPriceForm(TypeOfOperation, Model)
                Case CRUDOperations.CRUD_OPERATION_DELETE
                    RaiseEvent DeleteRecordFromPriceForm(TypeOfOperation, Model)
                Case CRUDOperations.CRUD_OPERATION_REJECT
                    RaiseEvent RejectRecordFromPriceForm(TypeOfOperation, Model)
                Case CRUDOperations.CRUD_OPERATION_APPROVE
                    RaiseEvent ApproveRecordFromPriceForm(TypeOfOperation, Model)
            End Select
        Else
            RaiseEvent OnSomethingWentWrong(response, TYPE_CRITICAL)
        End If
    End With
End Sub

Private Sub MainView_ClosePriceFormFrame()
    RaiseEvent ClosePriceForm
End Sub

Private Sub MainView_ResetPriceFormFrame()
    RaiseEvent ResetPriceForm
End Sub

'-------------------------------------------------------------------------
'Class Event
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Model = Nothing
    Set MainView = Nothing
End Sub
