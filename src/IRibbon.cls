VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRibbon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@IgnoreModule ImplementedInterfaceMember
'@Folder("System.Ribbon")
'@Interface
'@Exposed
Option Explicit

Public Sub Invalidate()
End Sub

Public Sub OnAction(ByVal Control As Office.IRibbonControl)
End Sub

Public Function GetImage(ByVal Identifier As String, ByRef outImage As Variant) As Variant
End Function

Public Function GetLabel(ByVal Identifier As String, ByRef outLabel As Variant) As Variant
End Function

Public Function GetScreenTip(ByVal Identifier As String, ByRef outTipp As Variant) As Variant
End Function

Public Function GetSize(ByVal Identifier As String, ByRef outSize As Variant) As Variant
End Function

Public Function IsVisible(ByVal Control As Office.IRibbonControl, ByRef outBoolean As Variant) As Variant
End Function

