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
'@Folder("MVP.ViewResizer")
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

Private This As TResizeView

Public Property Get View() As Object
    Guard.DefaultInstance Me
    Set View = This.View
End Property

Public Property Set View(ByVal RHS As Object)
    Guard.DefaultInstance Me
    Guard.NullReference RHS
    Set This.View = RHS
End Property

Public Property Get MinWidth() As Double
    MinWidth = This.MinWidth
End Property

Public Property Let MinWidth(ByVal RHS As Double)
    This.MinWidth = RHS
End Property

Public Property Get MinHeight() As Double
    MinHeight = This.MinHeight
End Property

Public Property Let MinHeight(ByVal RHS As Double)
    This.MinHeight = RHS
End Property

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is ResizeView
End Property

Private Sub Class_Initialize()
    If Not IsDefaultInstance Then
        Set This.LayoutBindings = New VBA.Collection
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
    View.Width = This.MinWidth
    View.Height = This.MinHeight
End Sub

Private Sub BindControlLayout(ByVal FormControl As MSForms.Control, ByVal Anchor As AnchorEdges)
    Guard.NullReference FormControl

    Dim Layout As ControlLayout
    Set Layout = New ControlLayout
    Layout.Bind This.View, FormControl, Anchor
    
    This.LayoutBindings.Add Layout
End Sub

Public Function IsViewResizable(ByVal View As Object, ByRef isMinimized As Boolean) As Boolean
    Guard.NullReference View
    
    On Error Resume Next
    If View.Width < This.MinWidth Then
        View.Width = This.MinWidth
        If Err.Number = 384 Then isMinimized = True: Exit Function 'Minimizeing View, resize runtime error
        IsViewResizable = False: Exit Function
    End If
    
    If View.Height < This.MinHeight Then
        View.Height = This.MinHeight
        If Err.Number = 384 Then isMinimized = True: Exit Function 'Minimizeing View, resize runtime error
        IsViewResizable = False: Exit Function
    End If
    
    On Error GoTo CleanFail
    Dim Layout As ControlLayout
    For Each Layout In This.LayoutBindings
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

    If This.Disposed Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If
    
    If Not This.LayoutBindings Is Nothing Then
        This.Disposed = Disposable.TryDispose(This.LayoutBindings)

        '@Ignore VariableNotUsed
        Dim i As Long
        For i = 1 To This.LayoutBindings.Count
            Disposable.TryDispose This.LayoutBindings.Item(This.LayoutBindings.Count)
            This.LayoutBindings.Remove This.LayoutBindings.Count
        Next i
    
        Set This.LayoutBindings = Nothing
    End If
    
    This.Disposed = True
    
    #If TestMode Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " is terminating"
    #End If
    
End Sub

Private Sub Class_Terminate()
    If Not This.Disposed Then Dispose
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

