VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BindingPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@Folder("Structure.Bindings")
'@PredeclaredId
'Copyright (c) 2020 Mathieu Guindon
Option Explicit

Implements IBindingPath

Private Type TState
    Context As Object
    Path As String
    Object As Object
    PropertyName As String
End Type

Private this As TState

'@Description "Gets/Sets the binding context."
Public Property Get Context() As Object
Attribute Context.VB_Description = "Gets/Sets the binding context."
    Set Context = this.Context
End Property

Public Property Set Context(ByVal RHS As Object)
    GuardClauses.GuardDefaultInstance Me, BindingPath, VBA.Information.TypeName(Me)
    GuardClauses.GuardNullReference RHS, VBA.Information.TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.Context, VBA.Information.TypeName(Me)
    Set this.Context = RHS
End Property

'@Description "Gets/Sets a string representing a property path against the binding context."
Public Property Get Path() As String
Attribute Path.VB_Description = "Gets/Sets a string representing a property path against the binding context."
    Path = this.Path
End Property

Public Property Let Path(ByVal RHS As String)
    GuardClauses.GuardDefaultInstance Me, BindingPath, VBA.Information.TypeName(Me)
    GuardClauses.GuardEmptyString RHS, VBA.Information.TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.Path, VBA.Information.TypeName(Me)
    this.Path = RHS
End Property

'@Description "Gets the bound object reference."
Public Property Get Object() As Object
Attribute Object.VB_Description = "Gets the bound object reference."
    Set Object = this.Object
End Property

'@Description "Gets the name of the bound property."
Public Property Get PropertyName() As String
Attribute PropertyName.VB_Description = "Gets the name of the bound property."
    PropertyName = this.PropertyName
End Property

'@Description "Resolves the Path to a bound object and property."
Public Sub Resolve()
Attribute Resolve.VB_Description = "Resolves the Path to a bound object and property."
    this.PropertyName = ResolvePropertyName(this.Path)
    Set this.Object = ResolvePropertyPath(this.Context, this.Path)
End Sub

'@Description "Creates a new binding path from the specified property path string and binding context."
Public Function Create(ByVal Context As Object, ByVal Path As String) As IBindingPath
Attribute Create.VB_Description = "Creates a new binding path from the specified property path string and binding context."
    GuardClauses.GuardNonDefaultInstance Me, BindingPath, VBA.Information.TypeName(Me)
    GuardClauses.GuardNullReference Context, VBA.Information.TypeName(Me)
    GuardClauses.GuardEmptyString Path, VBA.Information.TypeName(Me)
    
    Dim result As BindingPath
    Set result = New BindingPath
    Set result.Context = Context
    
    result.Path = Path
    result.Resolve
    
    Set Create = result
    
End Function

Private Function ResolvePropertyName(ByVal PropertyPath As String) As String
    Dim Parts As Variant
    Parts = VBA.Strings.Split(PropertyPath, ".")
    ResolvePropertyName = Parts(UBound(Parts))
End Function

Private Function ResolvePropertyPath(ByVal Context As Object, ByVal PropertyPath As String) As Object

    Dim Parts As Variant
    Parts = VBA.Strings.Split(PropertyPath, ".")
    
    If UBound(Parts) = LBound(Parts) Then
        Set ResolvePropertyPath = Context
    Else
        Dim RecursiveProperty As Object
        Set RecursiveProperty = VBA.Interaction.CallByName(Context, Parts(0), VBA.VbCallType.VbGet)
        If RecursiveProperty Is Nothing Then Exit Function
        Set ResolvePropertyPath = ResolvePropertyPath(RecursiveProperty, VBA.Strings.Right$(PropertyPath, VBA.Strings.Len(PropertyPath) - VBA.Strings.Len(Parts(0)) - 1))
    End If
    
End Function

Private Property Get IBindingPath_Context() As Object
    Set IBindingPath_Context = this.Context
End Property

Private Property Get IBindingPath_Path() As String
    IBindingPath_Path = this.Path
End Property

Private Property Get IBindingPath_Object() As Object
    Set IBindingPath_Object = this.Object
End Property

Private Property Get IBindingPath_PropertyName() As String
    IBindingPath_PropertyName = this.PropertyName
End Property

Private Sub IBindingPath_Resolve()
    Set this.Object = ResolvePropertyPath(this.Context, this.Path)
End Sub

Private Function IBindingPath_TryReadPropertyValue(ByRef outValue As Variant) As Boolean
    If this.Object Is Nothing Then Resolve
    On Error Resume Next
    outValue = VBA.Interaction.CallByName(this.Object, this.PropertyName, VBA.VbCallType.VbGet)
    IBindingPath_TryReadPropertyValue = (VBA.Information.Err.Number = 0)
    On Error GoTo 0
End Function

Private Function IBindingPath_TryWritePropertyValue(ByVal Value As Variant) As Boolean
    If this.Object Is Nothing Then Resolve
    On Error Resume Next
    VBA.Interaction.CallByName this.Object, this.PropertyName, VBA.VbCallType.VbLet, Value
    IBindingPath_TryWritePropertyValue = (VBA.Information.Err.Number = 0)
    On Error GoTo 0
End Function

Private Function IBindingPath_ToString() As String
    IBindingPath_ToString = StringBuilder.AppendFormat("Context: {0}; Path: {1}", VBA.Information.TypeName(this.Context), this.Path).ToString
End Function
