VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("AppObjects.Repository")
'@Interface
'@Exposed
Option Explicit

'public properties

Public Enum RepositoryType
    TYPE_POSTGRESQL = 1
    TYPE_ACCESS = 2
End Enum

Public Property Get IsEverythingOkay() As Boolean
End Property

Public Property Let IsEverythingOkay(ByVal vNewValue As Boolean)
End Property

'public methods
Public Sub InItRepository(ByVal ContextTableName As String, _
                            Optional ContextConnectionString As String = vbNullString, _
                            Optional ContextTableHeaders As Variant = vbNullString)
End Sub

Public Sub CheckTableStatus()
End Sub

Public Function GetRecords() As Variant
End Function

Public Function GetRecordsByCondition(ByVal ColumnIndex As Long, ByVal condition As Variant)
End Function

Public Sub AddRecord(ParamArray LineItems() As Variant)
End Sub

Public Sub UpdateRecord(ByVal TargetIndex As Long, ParamArray LineItems() As Variant)
End Sub

Public Sub DeleteRecord(ByVal TargetIndex As Long)
End Sub

