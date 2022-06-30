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

Public Enum HeightInPercent
    vbHeight45 = 45
    vbHeight50 = 50
    vbHeight55 = 55
End Enum

Public Enum WidthInPercent
    vbWidth45 = 45
    vbWidth50 = 50
    vbWidth55 = 55
End Enum

Public Sub InitializeResize()
End Sub

Public Sub BindControlLayout(ByVal FormControl As MSForms.Control, ByVal Anchor As AnchorEdges)
End Sub

Public Sub SetResolutionPercent(ByVal View As Object)
End Sub

Public Function IsViewResizable(ByVal View As Object, ByRef isMinimized As Boolean) As Boolean
End Function
