VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyChangeNotifierBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Provides a base implementation for the INotifyPropertyChanged observer interface."
'@Folder("MVVM.Notifier")
'@ModuleDescription "Provides a base implementation for the INotifyPropertyChanged observer interface."
Option Explicit

Implements INotifyPropertyChanged
Implements IDisposable

Private Type TState
    Handlers As Collection
End Type
Private this As TState

Private Sub Class_Initialize()
    Set this.Handlers = New Collection
End Sub

Private Sub IDisposable_Dispose()
    Set this.Handlers = Nothing
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    Dim Handler As IHandlePropertyChanged
    For Each Handler In this.Handlers
        Handler.HandlePropertyChanged Source, PropertyName
    Next
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    this.Handlers.Add Handler
End Sub
