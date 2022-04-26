VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBindingPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An object that can resolve a string property path to a value."
'@ModuleDescription "An object that can resolve a string property path to a value."
'@Folder("Structure.Bindings")
'@Interface
'@Exposed
Option Explicit

'@Description "Gets the context for resolving the Path."
Public Property Get Context() As Object
End Property

'@Description "Gets the property path of the binding."
Public Property Get Path() As String
End Property

'@Description "Gets a reference to the binding source object resolved from the Path."
Public Property Get Object() As Object
End Property

'@Description "Gets the name of the bound property resolved from the Path."
Public Property Get PropertyName() As String
End Property

'@Description "Reads and returns the property value; False on error, True if successful."
Public Function TryReadPropertyValue(ByRef outValue As Variant) As Boolean
End Function

'@Description "Writes the value to the bound property; False on error, True if successful."
Public Function TryWritePropertyValue(ByVal Value As Variant) As Boolean
End Function

'@Description "Returns a string representing this binding path for debugging purposes."
Public Function ToString() As String
End Function

'@Description "Resolves the binding source."
Public Sub Resolve()
End Sub

