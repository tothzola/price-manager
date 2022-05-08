VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Services"
Option Explicit

'public properties

Public Property Get IsEverythingOkayInEngine() As Boolean
End Property

Public Property Let IsEverythingOkayInEngine(ByVal vNewValue As Boolean)
End Property

Public Property Set Repository(ByVal vNewValue As IRepository)
End Property

Public Property Let ContextTableName(ByVal vNewValue As String)
End Property

Public Property Let ContextTableHeaders(ByVal vNewValue As Variant)
End Property

Public Property Let contextLineItems(ByVal vNewValue As Variant)
End Property

'public methods for communication

Public Sub InItService(ByVal SelectedRepositoryType As RepositoryType, ByVal TableName As String, ByVal TableHeaders As Variant)
End Sub

Public Function GetListOfAllRecordsFromContextTable() As Variant
End Function

Public Function GetSpecificListOfRecordsFromContextTable(ByVal ColumnReference As Long, ByVal SearchCondition As Variant) As Variant
End Function

Public Sub AddNewRecordToContextTable(ParamArray contextLineItems() As Variant)
End Sub

Public Sub UpdateRecordOfContextTable(ByVal index As Long, ParamArray contextLineItems() As Variant)
End Sub

Public Sub DeleteRecordofContextTable(ByVal index As Long)
End Sub

