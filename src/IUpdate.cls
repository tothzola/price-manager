VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("System.Updater")
'@Interface
'@Exposed
Option Explicit

Public Property Get NoUpdate() As Boolean
End Property

Public Function Available() As Boolean
End Function

Public Sub Download()
End Sub
