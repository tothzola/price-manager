VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IContextModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object keeping property bindings in scope."
'@ModuleDescription "An object keeping property bindings in scope."
'@Folder("Structure")
'@Interface
'@Exposed
'Copyright (c) 2020 Mathieu Guindon
Option Explicit

'@Description "Gets an object responsible for managing property bindings."
Public Property Get Bindings() As IBindingPerformer
Attribute Bindings.VB_Description = "Gets an object responsible for managing property bindings."
End Property

'@Description "Gets an abstract factory that creates objects that can format strings."
Public Property Get StringFormatterFactory() As IStringFormatterFactory
Attribute StringFormatterFactory.VB_Description = "Gets an abstract factory that creates objects that can format strings."
End Property

'@Description "Gets Debugging option True or False"
Public Property Get DebugOutput() As Boolean
Attribute DebugOutput.VB_Description = "Gets Debugging option True or False"
End Property
