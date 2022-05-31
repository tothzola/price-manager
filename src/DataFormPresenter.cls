VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataFormPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "DataForm"
Option Explicit

'-------------------------------------------------------------------------

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MainView As PAM
Attribute MainView.VB_VarHelpID = -1

'-------------------------------------------------------------------------

Public Event OpenDataForm(ByVal DFModel As DataFormModel, ByVal ContainerIdentifier As DataContainer)
Public Event CloseDataForm()
Public Event EditRecordFromDataForm(ByVal DFModel As DataFormModel)
Public Event OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)
Public Event FilterAndSortDataFormList()
Public Event PopulateValuesComboBox(ByVal DFModel As DataFormModel, ByVal TargetColumnIndex As Long)
Public Event ResetDataForm(ByVal DFModel As DataFormModel, ByVal ContainerIdentifier As DataContainer)
'-------------------------------------------------------------------------

Private Type TDataFormComponents
    model As DataFormModel
End Type

Private this As TDataFormComponents

'-------------------------------------------------------------------------

Private Property Get model() As DataFormModel
    Set model = this.model
End Property

Private Property Set model(ByVal vNewValue As DataFormModel)
    Set this.model = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt This Presenter with necessary parameters / Constructure procedure
'-------------------------------------------------------------------------

Public Sub InItDataForm(ByVal view As PAM, ByVal ContainerIdentifier As DataContainer)
    If model Is Nothing Then Set model = New DataFormModel
    If MainView Is Nothing Then Set MainView = view
    RaiseEvent OpenDataForm(model, ContainerIdentifier)
End Sub

'-------------------------------------------------------------------------
'Events Procedure
'-------------------------------------------------------------------------

Private Sub MainView_CloseDataFormFrame()
    RaiseEvent CloseDataForm
End Sub

Private Sub MainView_EditRecordFromDataFormFrame()
    Dim response As Variant
    response = model.IsValidRecordToEdit
    If response = True Then
        RaiseEvent EditRecordFromDataForm(model)
    Else
        RaiseEvent OnSomethingWentWrong(response, TYPE_INFORMATION)
    End If
End Sub

Private Sub MainView_FilterAndSortListFromDataFormFrame()
    RaiseEvent FilterAndSortDataFormList
End Sub

Private Sub MainView_OpenDataFormFrame(ByVal ContainerIdentifier As DataContainer)
    RaiseEvent OpenDataForm(model, ContainerIdentifier)
End Sub

Private Sub MainView_PopulateValuesList(ByVal TargetColumn As String)
    RaiseEvent PopulateValuesComboBox(model, model.GetTargetColumnIndex(TargetColumn))
End Sub

Private Sub MainView_ResetDataFormFrame(ByVal ContainerIdentifier As DataContainer)
    RaiseEvent ResetDataForm(model, ContainerIdentifier)
End Sub

'-------------------------------------------------------------------------
'Class Event
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set model = Nothing
    Set MainView = Nothing
End Sub

