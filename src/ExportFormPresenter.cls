VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExportFormPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ExportForm"
Option Explicit

'-------------------------------------------------------------------------

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MainView As PAM
Attribute MainView.VB_VarHelpID = -1

'-------------------------------------------------------------------------

Public Event OpenExportForm(ByVal EFModel As ExportFormModel)
Public Event CloseExportForm()
Public Event ResetExportForm()
Public Event ExportRecordsFromExportForm(ByVal EFModel As ExportFormModel)
Public Event OnSomethingWentWrong(ByVal Message As String, ByVal typeOfMessage As messageType)

'-------------------------------------------------------------------------

Private Type TExportFormComponents
    model As ExportFormModel
End Type

Private this As TExportFormComponents

'-------------------------------------------------------------------------

Private Property Get model() As ExportFormModel
    Set model = this.model
End Property

Private Property Set model(ByVal vNewValue As ExportFormModel)
    Set this.model = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt This Presenter with necessary parameters / Constructure procedure
'-------------------------------------------------------------------------

Public Sub InItExportForm(ByVal view As PAM)
    If model Is Nothing Then Set model = New ExportFormModel
    If MainView Is Nothing Then Set MainView = view
    RaiseEvent OpenExportForm(model)
End Sub

'-------------------------------------------------------------------------
'Events Procedure
'-------------------------------------------------------------------------

Private Sub MainView_CloseExportFormFrame()
    RaiseEvent CloseExportForm
End Sub

Private Sub MainView_ExportReport()
    Dim response As Variant
    response = model.IsEverythingOkayToExport
    If response = True Then
        Call model.GenerateExportData
        RaiseEvent ExportRecordsFromExportForm(model)
    Else
        RaiseEvent OnSomethingWentWrong(response, TYPE_INFORMATION)
    End If
End Sub

Private Sub MainView_OpenExportFormFrame()
    RaiseEvent OpenExportForm(model)
End Sub

Private Sub MainView_ResetExportFormFrame()
    RaiseEvent ResetExportForm
End Sub


'-------------------------------------------------------------------------
'Class Event
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set model = Nothing
    Set MainView = Nothing
End Sub

