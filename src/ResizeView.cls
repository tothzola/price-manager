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
'@Folder("AppObjects.ViewResizer")
'@PredeclaredId
'@Exposed
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" ( _
    ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function GetDC Lib "user32" ( _
    ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" ( _
    ByVal hDC As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function ReleaseDC Lib "user32" ( _
    ByVal hWnd As Long, ByVal hDC As LongPtr) As Long
#Else
    Private Declare Function GetSystemMetrics32 Lib "user32" Alias "GetSystemMetrics" ( _
                            ByVal nIndex As Long) As Long
    Private Declare Function GetDC Lib "user32" ( _
                            ByVal hWnd As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" ( _
                            ByVal hDC As Long, ByVal nIndex As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" ( _
                            ByVal hWnd As Long, ByVal hDC As Long) As Long
#End If

Implements IResizeView
Implements IDisposable

Private Const LOGPIXELSX As Long = 88             'Pixels/inch in X
Private Const POINTS_PER_INCH As Long = 72        'A point is defined as 1/72 inches

Private Type TResizeView
    View As Object
    LayoutBindings As VBA.Collection
    
    MinWidth As Double
    MinHeight As Double
    PercentHeight As Long
    PercentWidth As Long

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

Public Property Get LayoutBindings() As VBA.Collection
    Guard.DefaultInstance Me
    Set LayoutBindings = This.LayoutBindings
End Property

Public Property Set LayoutBindings(ByVal RHS As VBA.Collection)
    Guard.DefaultInstance Me
    Guard.NullReference RHS
    Set This.LayoutBindings = RHS
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

Public Property Get PercentHeight() As Long
    PercentHeight = This.PercentHeight
End Property

Public Property Let PercentHeight(ByVal RHS As Long)
    This.PercentHeight = RHS
End Property

Public Property Get PercentWidth() As Long
    PercentWidth = This.PercentWidth
End Property

Public Property Let PercentWidth(ByVal RHS As Long)
    This.PercentWidth = RHS
End Property

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is ResizeView
End Property

Private Sub Class_Initialize()
    If Not IsDefaultInstance Then
        Set LayoutBindings = New VBA.Collection
    End If
End Sub

'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As ResizeView
    Set Class = ResizeView
End Property

Public Function Create(ByVal View As Object, ByVal PercentHeight As WindowHeightPercent, ByVal PercentWidth As WindowWidthPercent) As IResizeView
    Guard.NonDefaultInstance Me
    
    Dim result As ResizeView
    Set result = New ResizeView
    Set result.View = View
    result.PercentHeight = PercentHeight
    result.PercentWidth = PercentWidth
    
    Set Create = result
End Function

Private Sub InitializeResize()
    ResizeViewResolutionPercent
End Sub

Private Sub SetResolutionPercent(ByVal View As Object)
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
    If View.Width < This.MinWidth - 5 Then
        View.Width = This.MinWidth
        If Err.Number = 384 Then isMinimized = True: Exit Function 'Minimizeing View, resize runtime error
        IsViewResizable = False
        Exit Function
    End If
    
    On Error GoTo CleanFail
    If View.Height < This.MinHeight - 5 Then
        View.Height = This.MinHeight
        IsViewResizable = False
        Exit Function
    End If
    
    Dim Layout As ControlLayout
    For Each Layout In LayoutBindings
        Layout.Resize View
    Next
    IsViewResizable = True
    
CleanExit:
    Exit Function

CleanFail:
    MsgBox Err.Number & vbTab & Err.Description, vbCritical, Title:=SIGN
    LogManager.Log ErrorLevel, "Error: " & VBA.Err.Number & ". " & VBA.Err.Description
    IsViewResizable = False
    Resume CleanExit
    Resume
    
End Function

'@Description "Get Points Per Pixel Screen resloution."
Private Function PointsPerPixel() As Double
Attribute PointsPerPixel.VB_Description = "Get Points Per Pixel Screen resloution."

    #If VBA7 Then
        '@Ignore UseMeaningfulName
        Dim hDC As LongPtr
        '@Ignore HungarianNotation
        Dim lDotsPerInch As LongPtr
    #Else
        Dim hDC As Long
        Dim lDotsPerInch As Long
    #End If

    hDC = GetDC(0)
    lDotsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
    PointsPerPixel = POINTS_PER_INCH / lDotsPerInch
    ReleaseDC 0, hDC

End Function

Private Sub ResizeViewResolutionPercent()
    MinWidth = GetSystemMetrics32(0) * PointsPerPixel * (PercentWidth / 100) 'UF Width in Resolution * DPI * PercentWidth
    MinHeight = GetSystemMetrics32(1) * PointsPerPixel * (PercentHeight / 100) 'UF Height in Resolution * DPI * PercentHeight
End Sub

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

Private Sub IResizeView_InitializeResize()
    InitializeResize
End Sub

Private Sub IResizeView_SetResolutionPercent(ByVal View As Object)
    SetResolutionPercent View
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
