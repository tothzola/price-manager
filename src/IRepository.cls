VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Repository"
Option Explicit

Public Function GetRecords(ByVal TargetNamedRange As String) As Variant
End Function

Public Sub AddRecord(ByVal TargetNamedRange As String, ParamArray LineItems() As Variant)
End Sub

Public Sub UpdateRecord(ByVal TargetNamedRange As String, ByVal TargetIndex As Long, ParamArray LineItems() As Variant)
End Sub

Public Sub DeleteRecord(ByVal TargetNamedRange As String, ByVal TargetIndex As Long)
End Sub
