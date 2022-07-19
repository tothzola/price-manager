VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("PriceApproval.Services")
'@Interface
'@Exposed
Option Explicit

'public properties
Public Property Get IsEverythingOkayInEngine() As Boolean
End Property

Public Property Let IsEverythingOkayInEngine(ByVal vNewValue As Boolean)
End Property

'@Description "public methods for communication"
Public Sub InItService(ByVal SelectedRepositoryType As RepositoryType, ByVal TableName As String, ByVal TableHeaders As Variant, _
Optional ByVal ContextConnectionString As String = vbNullString)
Attribute InItService.VB_Description = "public methods for communication"

End Sub

Public Function GetListOfAllRecordsFromContextTable() As Variant
End Function

Public Function GetSpecificListOfRecordsFromContextTable(ByVal ColumnReference As Long, ByVal SearchCondition As Variant) As Variant
End Function

Public Sub AddNewRecordToContextTable(ParamArray ContextLineItems() As Variant)
End Sub

Public Sub UpdateRecordOfContextTable(ByVal Index As Long, ParamArray ContextLineItems() As Variant)
End Sub

Public Sub DeleteRecordofContextTable(ByVal Index As Long)
End Sub
