VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResizeView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("MVVM.ViewResizer")
'@PredeclaredId
'@Exposed
Option Explicit

Implements IResizeView
Implements IDisposable

Private Type TResizeView
    View As Object
    LayoutBindings As VBA.Collection
    
    MinWidth As Double
    MinHeight As Double

    Disposed As Boolean
End Type

Private this As TResizeView

Public Property Get View() As Object
    Guard.DefaultInstance Me
    Set View = this.View
End Property

Public Property Set View(ByVal RHS As Object)
    Guard.DefaultInstance Me
    Guard.NullReference RHS
    Set this.View = RHS
End Property

Public Property Get MinWidth() As Double
    MinWidth = this.MinWidth
End Property

Public Property Let MinWidth(ByVal RHS As Double)
    this.MinWidth = RHS
End Property

Public Property Get MinHeight() As Double
    MinHeight = this.MinHeight
End Property

Public Property Let MinHeight(ByVal RHS As Double)
    this.MinHeight = RHS
End Property

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is ResizeView
End Property

Private Sub Class_Initialize()
    If Not IsDefaultInstance Then
        Set this.LayoutBindings = New VBA.Collection
    End If
End Sub

'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As ResizeView
    Set Class = ResizeView
End Property

Public Function Create(ByVal View As Object, ByVal DefaultHeight As Double, ByVal DefaultWidth As Double) As IResizeView
    Guard.NonDefaultInstance Me
    
    Dim result As ResizeView
    Set result = New ResizeView
    Set result.View = View
    result.MinHeight = DefaultHeight
    result.MinWidth = DefaultWidth
    
    Set Create = result
End Function

Private Sub SetDefaultSize(ByVal View As Object)
    View.Width = this.MinWidth
    View.Height = this.MinHeight
End Sub

Private Sub BindControlLayout(ByVal FormControl As MSForms.Control, ByVal Anchor As AnchorEdges)
    Guard.NullReference FormControl

    Dim Layout As ControlLayout
    Set Layout = New ControlLayout
    Layout.Bind this.View, FormControl, Anchor
    
    this.LayoutBindings.Add Layout
End Sub

Public Function IsViewResizable(ByVal View As Object, ByRef isMinimized As Boolean) As Boolean
    Guard.NullReference View
    
    On Error Resume Next
    If View.Width < this.MinWidth Then
        View.Width = this.MinWidth
        If Err.Number = 384 Then isMinimized = True: Exit Function 'Minimizeing View, resize runtime error
        IsViewResizable = False: Exit Function
    End If
    
    If View.Height < this.MinHeight Then
        View.Height = this.MinHeight
        If Err.Number = 384 Then isMinimized = True: Exit Function 'Minimizeing View, resize runtime error
        IsViewResizable = False: Exit Function
    End If
    
    On Error GoTo CleanFail
    Dim Layout As ControlLayout
    For Each Layout In this.LayoutBindings
        Layout.Resize View
    Next
    
CleanExit:
    IsViewResizable = True
    Exit Function

CleanFail:
    MsgBox Err.Number & vbTab & Err.Description, vbCritical, Title:=SIGN
    LogManager.Log ErrorLevel, "Error: " & VBA.Err.Number & ". " & VBA.Err.Description
    IsViewResizable = False
    Resume CleanExit
    Resume
    
End Function

Private Sub Dispose()

    If this.Disposed Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If
    
    If Not this.LayoutBindings Is Nothing Then
        this.Disposed = Disposable.TryDispose(this.LayoutBindings)

        '@Ignore VariableNotUsed
        Dim i As Long
        For i = 1 To this.LayoutBindings.Count
            Disposable.TryDispose this.LayoutBindings.Item(this.LayoutBindings.Count)
            this.LayoutBindings.Remove this.LayoutBindings.Count
        Next i
    
        Set this.LayoutBindings = Nothing
    End If
    
    this.Disposed = True
    
    #If TestMode Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " is terminating"
    #End If
    
End Sub

Private Sub Class_Terminate()
    If Not this.Disposed Then Dispose
End Sub

Private Sub IResizeView_SetDefaultSize(ByVal View As Object)
    SetDefaultSize View
End Sub

Private Sub IResizeView_BindControlLayout(ByVal FormControl As MSForms.Control, ByVal Anchor As AnchorEdges)
    BindControlLayout FormControl, Anchor
End Sub

Private Function IResizeView_IsViewResizable(ByVal View As Object, ByRef isMinimized As Boolean) As Boolean
    IResizeView_IsViewResizable = IsViewResizable(View, isMinimized)
End Function

Private Sub IDisposable_Dispose()
    Dispose
End Sub

