VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISharePointDOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("System.Update")
'@Interface

Public Function IsConnected() As Boolean
End Function

Public Function GetAddInVersion(ByVal Name As String, ByRef outVersion As String) As Boolean
End Function

