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
'@Folder("MVVM.ViewResizer.Control")
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

Private this As TControlLayout

Public Sub Bind(ByVal form As Object, ByVal FormControl As MSForms.Control, ByVal Anchor As AnchorEdges)

    Set this.FormSize = New ControlSize
    Set this.Size = New ControlSize
    Set this.Margins = New SideMargins
    this.Anchors = Anchor

    this.FormSize.Height = form.InsideHeight
    this.FormSize.Width = form.InsideWidth

    this.Size.Height = FormControl.Height
    this.Size.Width = FormControl.Width

    If (Anchor And BottomAnchor) = BottomAnchor Then
        this.Margins.BottomMargin = this.FormSize.Height - FormControl.Top - FormControl.Height
    End If

    If (Anchor And LeftAnchor) = LeftAnchor Then
        this.Margins.LeftMargin = FormControl.Left
    End If

    If (Anchor And RightAnchor) = RightAnchor Then
        this.Margins.RightMargin = this.FormSize.Width - FormControl.Left - FormControl.Width
    End If

    If (Anchor And TopAnchor) = TopAnchor Then
        this.Margins.TopMargin = FormControl.Top
    End If

    Set this.BoundControl = FormControl

End Sub

Public Sub Resize(ByVal form As Object)

    this.FormSize.Height = form.InsideHeight
    this.FormSize.Width = form.InsideWidth

    Layout this.BoundControl

End Sub

Private Sub Layout(ByVal FormControl As MSForms.Control)

    If (Anchors And TopAnchor) = TopAnchor Then

        If (Anchors And BottomAnchor) = BottomAnchor Then
            FormControl.Height = this.FormSize.Height - FormControl.Top - this.Margins.BottomMargin
        End If

    ElseIf (Anchors And BottomAnchor) = BottomAnchor Then
        FormControl.Top = this.FormSize.Height - FormControl.Height - this.Margins.BottomMargin
    End If


    If (Anchors And LeftAnchor) = LeftAnchor Then

        If (Anchors And RightAnchor) = RightAnchor Then
            FormControl.Width = this.FormSize.Width - FormControl.Left - this.Margins.RightMargin
        End If

    ElseIf (Anchors And RightAnchor) = RightAnchor Then
        FormControl.Left = this.FormSize.Width - FormControl.Width - this.Margins.RightMargin
    End If

End Sub

Public Property Get Anchors() As AnchorEdges
    Anchors = this.Anchors
End Property

Public Property Let Anchors(ByVal Ancor As AnchorEdges)
    this.Anchors = Ancor
End Property

Public Property Get BoundControl() As MSForms.Control
    Set BoundControl = this.BoundControl
End Property

Public Property Set BoundControl(ByVal bound As MSForms.Control)
    Set this.BoundControl = bound
End Property
