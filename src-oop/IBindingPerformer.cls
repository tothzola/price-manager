VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IBindingPerformer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object responsible for managing property bindings in an application context."
'@ModuleDescription "Represents an object responsible for managing property bindings in an application context."
'@Folder("Structure.Bindings")
'@Interface
'@Exposed
Option Explicit

'@Description "Creates a new property binding for the specified path, source, and target."
Public Function BindPropertyPath(ByVal Source As Object, ByVal PropertyPath As String, ByVal Target As Object, _
                                 Optional ByVal TargetProperty As String, _
                                 Optional ByVal Mode As BindingMode, _
                                 Optional ByVal UpdateTrigger As BindingUpdateSourceTrigger, _
                                 Optional ByVal StringFormat As String) As IPropertyBinding

End Function

'@Description "Gets a reference to an object that creates notification helper object instances."
Public Property Get NotifierFactory() As INotifierFactory
End Property

'@Description "Gets a reference to an object that creates string-formatting helper object instances."
Public Property Get StringFormatterFactory() As IStringFormatterFactory
End Property

'@Description "Gets a value indicating whether bindings produce debugger output."
Public Property Get DebugOutput() As Boolean
End Property

'@Description "Applies all property bindings."
Public Sub Apply(ByVal Context As Object)
End Sub


