VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExportFormPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVP.Presenter.ExportForm")
Option Explicit

'-------------------------------------------------------------------------

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents MainView As PriceApprovalView
Attribute MainView.VB_VarHelpID = -1

'-------------------------------------------------------------------------

Public Event OpenExportForm(ByVal EFModel As ExportFormModel)
Public Event CloseExportForm()
Public Event ResetExportForm()
Public Event ExportRecordsFromExportForm(ByVal EFModel As ExportFormModel)
Public Event OnSomethingWentWrong(ByVal message As String, ByVal typeOfMessage As messageType)

'-------------------------------------------------------------------------

Private Type TExportFormComponents
    Model As ExportFormModel
End Type

Private this As TExportFormComponents

'-------------------------------------------------------------------------

Private Property Get Model() As ExportFormModel
    Set Model = this.Model
End Property

Private Property Set Model(ByVal vNewValue As ExportFormModel)
    Set this.Model = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt This Presenter with necessary parameters / Constructure procedure
'-------------------------------------------------------------------------

Public Sub InItExportForm(ByVal View As PriceApprovalView)
    If Model Is Nothing Then Set Model = New ExportFormModel
    If MainView Is Nothing Then Set MainView = View
    RaiseEvent OpenExportForm(Model)
End Sub

'-------------------------------------------------------------------------
'Events Procedure
'-------------------------------------------------------------------------

Private Sub MainView_CloseExportFormFrame()
    RaiseEvent CloseExportForm
End Sub

Private Sub MainView_ExportReport()
    Dim response As Variant
    response = Model.IsEverythingOkayToExport
    If response = True Then
        Call Model.GenerateExportData
        RaiseEvent ExportRecordsFromExportForm(Model)
    Else
        RaiseEvent OnSomethingWentWrong(response, TYPE_INFORMATION)
    End If
End Sub

Private Sub MainView_OpenExportFormFrame()
    RaiseEvent OpenExportForm(Model)
End Sub

Private Sub MainView_ResetExportFormFrame()
    RaiseEvent ResetExportForm
End Sub


'-------------------------------------------------------------------------
'Class Event
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Model = Nothing
    Set MainView = Nothing
End Sub

