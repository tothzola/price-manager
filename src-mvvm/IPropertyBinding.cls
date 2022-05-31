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
'Copyright (c) 2020 Mathieu Guindon
Option Explicit

Public Enum ApplyResult
    BindingSkipped = 0
    BindingSuccess
    BindingFailed
End Enum

Public Enum BindingMode
    TwoWayBinding = 0
    OneWayBinding
    OneWayToSource
    OneTimeBinding
End Enum

Public Enum BindingUpdateSourceTrigger
    Never = 0
    OnChange
    OnExit
    OnKeyPress
End Enum

'@Description "Gets a value indicating the binding mode/direction."
Public Property Get Mode() As BindingMode
Attribute Mode.VB_Description = "Gets a value indicating the binding mode/direction."
End Property

'@Description "Gets a value indicating the binding update trigger."
Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
Attribute UpdateSourceTrigger.VB_Description = "Gets a value indicating the binding update trigger."
End Property

'@Description "Gets the binding path resolving to the binding source."
Public Property Get Source() As IBindingPath
Attribute Source.VB_Description = "Gets the binding path resolving to the binding source."
End Property

'@Description "Gets the binding path resolving to the binding target."
Public Property Get Target() As IBindingPath
Attribute Target.VB_Description = "Gets the binding path resolving to the binding target."
End Property

'@Description "Gets the name of the target property implicitly bound for the Target class type."
Public Property Get DefaultTargetProperty() As String
Attribute DefaultTargetProperty.VB_Description = "Gets the name of the target property implicitly bound for the Target class type."
End Property

'@Description "Gets the string formatter (if any) used when applying the binding to the target."
Public Property Get StringFormat() As IStringFormatter
Attribute StringFormat.VB_Description = "Gets the string formatter (if any) used when applying the binding to the target."
End Property

'@Description "Applies the binding."
Public Sub Apply()
Attribute Apply.VB_Description = "Applies the binding."
End Sub
