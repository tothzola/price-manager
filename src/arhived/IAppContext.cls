VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@Folder("MVP.Model")
'@Exposed
'@Interface

Public Property Get MainService() As IServices
End Property

Public Property Get UserService() As IServices
End Property

Public Function IsRepositoryReachable() As Boolean
End Function
