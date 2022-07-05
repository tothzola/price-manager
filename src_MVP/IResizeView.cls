VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IResizeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@Folder("AppObjects.ViewResizer")
'@Exposed
'@Interface

Public Sub SetDefaultSize(ByVal View As Object)
End Sub

Public Sub BindControlLayout(ByVal FormControl As MSForms.Control, ByVal Anchor As AnchorEdges)
End Sub

Public Function IsViewResizable(ByVal View As Object, ByRef isMinimized As Boolean) As Boolean
End Function
