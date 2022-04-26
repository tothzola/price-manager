VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for binding a ViewModel property path to a UI element."
'@ModuleDescription "An object responsible for binding a ViewModel property path to a UI element."
'@Folder("Structure.Bindings")
'@Interface
'@Exposed
Option Explicit

Public Enum ApplyResult
    BindingSkipped = 0
    BindingSuccess = 1
    BindingFailed = 2
End Enum

Public Enum BindingMode
    TwoWayBinding = 0
    OneWayBinding = 1
    OneWayToSource = 2
    OneTimeBinding = 3
End Enum

Public Enum BindingUpdateSourceTrigger
    Never = 0
    OnPropertyChanged = 1
    OnEnter = 2
    OnExit = 3
    OnChange = 4
End Enum

'@Description "Gets a value indicating the binding mode/direction."
Public Property Get Mode() As BindingMode
End Property

'@Description "Gets a value indicating the binding update trigger."
Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
End Property

'@Description "Gets the binding path resolving to the binding source."
Public Property Get Source() As IBindingPath
End Property

'@Description "Gets the binding path resolving to the binding target."
Public Property Get Target() As IBindingPath
End Property

'@Description "Gets the name of the target property implicitly bound for the Target class type."
Public Property Get DefaultTargetProperty() As String
End Property

'@Description "Gets the string formatter (if any) used when applying the binding to the target."
Public Property Get StringFormat() As IStringFormatter
End Property

'@Description "Applies the binding."
Public Sub Apply()
End Sub

