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

Public Enum WindowHeightPercent
    vbHeightPercent20 = 20
    vbHeightPercent30 = 30
    vbHeightPercent40 = 40
    vbHeightPercent50 = 50
    vbHeightPercent60 = 60
    vbHeightPercent70 = 70
    vbHeightPercent80 = 80
End Enum

Public Enum WindowWidthPercent
    vbWidthPercent20 = 20
    vbWidthPercent30 = 30
    vbWidthPercent40 = 40
    vbWidthPercent50 = 50
    vbWidthPercent60 = 60
    vbWidthPercent70 = 70
    vbWidthPercent80 = 80
End Enum

Public Sub InitializeResize()
End Sub

Public Sub SetResolutionPercent(ByVal View As Object)
End Sub

Public Sub BindControlLayout(ByVal FormControl As MSForms.Control, ByVal Anchor As AnchorEdges)
End Sub

Public Function IsViewResizable(ByVal View As Object, ByRef isMinimized As Boolean) As Boolean
End Function
