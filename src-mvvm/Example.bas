Attribute VB_Name = "Example"
'@Folder("Example")
Option Explicit

'@Description "Runs the StringFormat example UI."
Public Sub Run()

    Dim ViewModel As ExampleViewModel
    Set ViewModel = ExampleViewModel.Create
    
    ViewModel.Title = "This is a Binding example for the Labels"
    ViewModel.FormatToGeneral = "Ex.: General Format"
    ViewModel.FormatToCustom = "Ex.: German Format"
    
    Dim Model As IContextModel
    Set Model = ContextModel.Create(DebugOutput:=True)
    
    Dim View As IView
    Set View = ExampleView.Create(Model, ViewModel)

    If View.ShowDialog Then
        Debug.Print ViewModel.Title
    Else
        Debug.Print "Dialog was cancelled."
    End If
    
    Disposable.TryDispose Model
    
End Sub

