VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExampleView 
   Caption         =   "ExampleView"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5625
   OleObjectBlob   =   "ExampleView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExampleView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An example implementation of a View."

'@ModuleDescription "An example implementation of a View."
'@Folder("Example")
Option Explicit

Implements IView
Implements ICancellable

Private Type TView
    Context As IContextModel
    ViewModel As ExampleViewModel
    IsCancelled As Boolean
End Type

Private this As TView

'@Description "A factory method to create new instances of this View, already wired-up to a ViewModel."
Public Function Create(ByVal Context As IContextModel, ByVal ViewModel As ExampleViewModel) As IView

    Dim result As ExampleView
    Set result = New ExampleView
    Set result.Context = Context
    Set result.ViewModel = ViewModel
    Set Create = result
    
End Function

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is ExampleView
End Property

'@Description "Gets/sets the ViewModel to use as a context for property and command bindings."
Public Property Get ViewModel() As ExampleViewModel
    Set ViewModel = this.ViewModel
End Property

Public Property Set ViewModel(ByVal RHS As ExampleViewModel)
    Set this.ViewModel = RHS
End Property

'@Description "Gets/sets the context model"
Public Property Get Context() As IContextModel
    Set Context = this.Context
End Property

Public Property Set Context(ByVal RHS As IContextModel)
    Set this.Context = RHS
End Property

Private Sub BindViewModelProperties()

    With Context.Bindings
        .BindPropertyPath ViewModel, "Title", Me.FrameInstructions
        .BindPropertyPath ViewModel, "FormatToGeneral", Me.LabelGeneral
        .BindPropertyPath ViewModel, "FormatToCustom", Me.LabelCustom

    End With
End Sub

Private Sub InitializeBindings()

    If ViewModel Is Nothing Then Exit Sub
    BindViewModelProperties
    this.Context.Bindings.Apply ViewModel
    
End Sub

Private Sub OnCancel()
    this.IsCancelled = True
    Me.Hide
End Sub

Private Property Get ICancellable_IsCancelled() As Boolean
    ICancellable_IsCancelled = this.IsCancelled
End Property

Private Sub ICancellable_OnCancel()
    OnCancel
End Sub

Private Sub IView_Hide()
    Me.Hide
End Sub

Private Sub IView_Show()
    InitializeBindings
    Me.Show vbModal
End Sub

Private Function IView_ShowDialog() As Boolean
    InitializeBindings
    Me.Show vbModal
    IView_ShowDialog = Not this.IsCancelled
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

