VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlLayout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Encapsulates basic layout logic for dynamic MSForms controls."
'@Folder("AppObjects.ViewResizer.Control")
'@ModuleDescription "Encapsulates basic layout logic for dynamic MSForms controls."
'@Exposed
Option Explicit

Public Enum AnchorEdges
    LeftAnchor = 1
    TopAnchor = 2
    RightAnchor = 4
    BottomAnchor = 8
    AnchorAll = LeftAnchor + TopAnchor + RightAnchor + BottomAnchor
End Enum

Private Type TControlLayout
    Anchors As AnchorEdges
    Margins As SideMargins
    FormSize As ControlSize
    Size As ControlSize
    BoundControl As MSForms.Control
End Type

Private This As TControlLayout

Public Sub Bind(ByVal form As Object, ByVal FormControl As MSForms.Control, ByVal Anchor As AnchorEdges)

    Set This.FormSize = New ControlSize
    Set This.Size = New ControlSize
    Set This.Margins = New SideMargins
    This.Anchors = Anchor

    This.FormSize.Height = form.InsideHeight
    This.FormSize.Width = form.InsideWidth

    This.Size.Height = FormControl.Height
    This.Size.Width = FormControl.Width

    If (Anchor And BottomAnchor) = BottomAnchor Then
        This.Margins.BottomMargin = This.FormSize.Height - FormControl.Top - FormControl.Height
    End If

    If (Anchor And LeftAnchor) = LeftAnchor Then
        This.Margins.LeftMargin = FormControl.Left
    End If

    If (Anchor And RightAnchor) = RightAnchor Then
        This.Margins.RightMargin = This.FormSize.Width - FormControl.Left - FormControl.Width
    End If

    If (Anchor And TopAnchor) = TopAnchor Then
        This.Margins.TopMargin = FormControl.Top
    End If

    Set This.BoundControl = FormControl

End Sub

Public Sub Resize(ByVal form As Object)

    This.FormSize.Height = form.InsideHeight
    This.FormSize.Width = form.InsideWidth

    Layout This.BoundControl

End Sub

Private Sub Layout(ByVal FormControl As MSForms.Control)

    If (Anchors And TopAnchor) = TopAnchor Then

        If (Anchors And BottomAnchor) = BottomAnchor Then
            FormControl.Height = This.FormSize.Height - FormControl.Top - This.Margins.BottomMargin
        End If

    ElseIf (Anchors And BottomAnchor) = BottomAnchor Then
        FormControl.Top = This.FormSize.Height - FormControl.Height - This.Margins.BottomMargin
    End If


    If (Anchors And LeftAnchor) = LeftAnchor Then

        If (Anchors And RightAnchor) = RightAnchor Then
            FormControl.Width = This.FormSize.Width - FormControl.Left - This.Margins.RightMargin
        End If

    ElseIf (Anchors And RightAnchor) = RightAnchor Then
        FormControl.Left = This.FormSize.Width - FormControl.Width - This.Margins.RightMargin
    End If

End Sub

Public Property Get Anchors() As AnchorEdges
    Anchors = This.Anchors
End Property

Public Property Let Anchors(ByVal Ancor As AnchorEdges)
    This.Anchors = Ancor
End Property

Public Property Get BoundControl() As MSForms.Control
    Set BoundControl = This.BoundControl
End Property

Public Property Set BoundControl(ByVal bound As MSForms.Control)
    Set This.BoundControl = bound
End Property
