VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataFormPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.Presenter.DataForm")
Option Explicit

'-------------------------------------------------------------------------

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MainView As PriceApprovalView
Attribute MainView.VB_VarHelpID = -1

'-------------------------------------------------------------------------

Public Event OpenDataForm(ByVal DFModel As DataFormModel, ByVal ContainerIdentifier As DataContainer)
Public Event CloseDataForm()
Public Event EditRecordFromDataForm(ByVal DFModel As DataFormModel)
Public Event OnSomethingWentWrong(ByVal Message As String, ByVal typeOfMessage As messageType)
Public Event FilterAndSortDataFormList()
Public Event PopulateValuesComboBox(ByVal DFModel As DataFormModel, ByVal TargetColumnIndex As Long)
Public Event ResetDataForm(ByVal DFModel As DataFormModel, ByVal ContainerIdentifier As DataContainer)
'-------------------------------------------------------------------------

Private Type TDataFormComponents
    Model As DataFormModel
End Type

Private This As TDataFormComponents

'-------------------------------------------------------------------------

Private Property Get Model() As DataFormModel
    Set Model = This.Model
End Property

Private Property Set Model(ByVal vNewValue As DataFormModel)
    Set This.Model = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt This Presenter with necessary parameters / Constructure procedure
'-------------------------------------------------------------------------

Public Sub InItDataForm(ByVal View As PriceApprovalView, ByVal ContainerIdentifier As DataContainer)
    If Model Is Nothing Then Set Model = New DataFormModel
    If MainView Is Nothing Then Set MainView = View
    RaiseEvent OpenDataForm(Model, ContainerIdentifier)
End Sub

'-------------------------------------------------------------------------
'Events Procedure
'-------------------------------------------------------------------------

Private Sub MainView_CloseDataFormFrame()
    RaiseEvent CloseDataForm
End Sub

Private Sub MainView_EditRecordFromDataFormFrame()
    Dim response As Variant
    response = Model.IsValidRecordToEdit
    If response = True Then
        RaiseEvent EditRecordFromDataForm(Model)
    Else
        RaiseEvent OnSomethingWentWrong(response, TYPE_INFORMATION)
    End If
End Sub

Private Sub MainView_FilterAndSortListFromDataFormFrame()
    RaiseEvent FilterAndSortDataFormList
End Sub

Private Sub MainView_OpenDataFormFrame(ByVal ContainerIdentifier As DataContainer)
    RaiseEvent OpenDataForm(Model, ContainerIdentifier)
End Sub

Private Sub MainView_PopulateValuesList()
    RaiseEvent PopulateValuesComboBox(Model, Model.selectedColumn)
End Sub

Private Sub MainView_ResetDataFormFrame(ByVal ContainerIdentifier As DataContainer)
    RaiseEvent ResetDataForm(Model, ContainerIdentifier)
End Sub

'-------------------------------------------------------------------------
'Class Event
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Model = Nothing
    Set MainView = Nothing
End Sub

