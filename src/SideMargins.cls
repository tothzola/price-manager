VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SideMargins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("PriceApproval.ViewResizer.Control")
Option Explicit

Private Type TSideMargins
    BottomMargin As Double
    LeftMargin As Double
    RightMargin As Double
    TopMargin As Double
End Type

Private this As TSideMargins

Public Property Get BottomMargin() As Double
    BottomMargin = this.BottomMargin
End Property

Public Property Let BottomMargin(ByVal Bottom As Double)
    this.BottomMargin = Bottom
End Property

Public Property Get LeftMargin() As Double
    LeftMargin = this.LeftMargin
End Property

Public Property Let LeftMargin(ByVal Left As Double)
    this.LeftMargin = Left
End Property

Public Property Get RightMargin() As Double
    RightMargin = this.RightMargin
End Property

Public Property Let RightMargin(ByVal Right As Double)
    this.RightMargin = Right
End Property

Public Property Get TopMargin() As Double
    TopMargin = this.TopMargin
End Property

Public Property Let TopMargin(ByVal Top As Double)
    this.TopMargin = Top
End Property
