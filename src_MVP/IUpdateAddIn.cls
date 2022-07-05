VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUpdateAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("System.Update")
'@Interface
'@Exposed
Option Explicit

Public Sub IsThereAnUpdate(Optional ByRef GetAppUpdates As Boolean)
End Sub

Public Function GetLastUpdate() As Date
End Function

