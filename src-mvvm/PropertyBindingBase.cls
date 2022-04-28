VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PropertyBindingBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for binding a ViewModel property path to a UI element."
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@ModuleDescription "An object responsible for binding a ViewModel property path to a UI element."
'@Folder("Structure.Bindings")
'@PredeclaredId
'@Exposed
'Copyright (c) 2020 Mathieu Guindon
Option Explicit

Implements IPropertyBinding
Implements IDisposable

Private Const DefaultTargetControlProperty As String = "Value"

Private Type TState

    Context As IContextModel
    
    Source As IBindingPath
    Target As IBindingPath
    
    Applied As Boolean
    Applying As Boolean
    
    Handlers As VBA.Collection
    
    Mode As BindingMode
    
    UpdateSourceTrigger As BindingUpdateSourceTrigger
    
    StringFormat As IStringFormatter
    
    Disposed As Boolean
End Type

Private this As TState

Public Property Get Context() As IContextModel
    Set Context = this.Context
End Property

Friend Property Set Context(ByVal RHS As IContextModel)
    GuardClauses.GuardDoubleInitialization this.Context, VBA.Information.TypeName(Me)
    GuardClauses.GuardNullReference RHS, VBA.Information.TypeName(Me)
    Set this.Context = RHS
End Property

Public Property Get Source() As IBindingPath
    Set Source = this.Source
End Property

Friend Property Set Source(ByVal RHS As IBindingPath)
    GuardClauses.GuardDoubleInitialization this.Source, VBA.Information.TypeName(Me)
    GuardClauses.GuardNullReference RHS, VBA.Information.TypeName(Me)
    Set this.Source = RHS
End Property

Public Property Get Target() As IBindingPath
    Set Target = this.Target
End Property

Friend Property Set Target(ByVal RHS As IBindingPath)
    GuardClauses.GuardDoubleInitialization this.Target, VBA.Information.TypeName(Me)
    GuardClauses.GuardNullReference RHS, VBA.Information.TypeName(Me)
    Set this.Target = RHS
End Property

Public Property Get Mode() As BindingMode
    Mode = this.Mode
End Property

Public Property Let Mode(ByVal RHS As BindingMode)
    this.Mode = RHS
End Property

Public Property Get UpdateSourceTrigger() As BindingUpdateSourceTrigger
    UpdateSourceTrigger = this.UpdateSourceTrigger
End Property

Public Property Let UpdateSourceTrigger(ByVal RHS As BindingUpdateSourceTrigger)
    this.UpdateSourceTrigger = RHS
End Property

Public Property Get StringFormat() As IStringFormatter
    Set StringFormat = this.StringFormat
End Property

Friend Property Set StringFormat(ByVal RHS As IStringFormatter)
    GuardClauses.GuardDoubleInitialization this.StringFormat, VBA.Information.TypeName(Me)
    Set this.StringFormat = RHS
End Property

Public Property Get DefaultTargetProperty() As String
    DefaultTargetProperty = DefaultTargetControlProperty
End Property

Private Property Get CanApplyToTarget() As Boolean
    'One-way to source and already-applied one-time binding modes do not apply to target"
    CanApplyToTarget = Not this.Applying And this.Mode <> OneWayToSource And (this.Mode <> OneTimeBinding Or Not this.Applied)
End Property

Private Property Get CanApplyToSource() As Boolean
    'One-time and one-way binding modes do not write values to the binding source (ViewModel)."
    CanApplyToSource = Not this.Applying And this.Mode <> OneTimeBinding And this.Mode <> OneWayBinding
End Property

Private Sub Class_Initialize()
    If Not Me Is PropertyBindingBase Then
        Set this.Handlers = New VBA.Collection
    End If
End Sub

Public Function Create(ByVal Context As IContextModel, ByVal Source As IBindingPath, ByVal TargetContext As Object, _
                       Optional ByVal TargetPropertyPath As String = DefaultTargetControlProperty, _
                       Optional ByVal Mode As BindingMode, _
                       Optional ByVal UpdateSource As BindingUpdateSourceTrigger, _
                       Optional ByVal StringFormat As IStringFormatter) As IPropertyBinding
                       
    GuardClauses.GuardNonDefaultInstance Me, PropertyBindingBase, VBA.Information.TypeName(Me)

    Dim result As PropertyBindingBase
    Set result = New PropertyBindingBase
    
    Set result.Context = Context
    Set result.Source = Source
    Set result.Target = BindingPath.Create(TargetContext, TargetPropertyPath)
    
    result.Mode = Mode
    result.UpdateSourceTrigger = UpdateSource

    Set result.StringFormat = StringFormat
    Set Create = result
    
End Function

Private Sub DebugMessage(ByVal Message As String)
    If this.Context.Bindings.DebugOutput Then
        Debug.Print "PropertyBinding: [" & ToString & "] " & Message
    End If
End Sub

Private Function ToString() As String
    ToString = VBA.Information.TypeName(this.Source.Context) & "." & _
                                                             this.Source.Path & " -> " & _
                                                             VBA.Information.TypeName(this.Target.Context) & "." & _
                                                             this.Target.Path
End Function

Public Sub Apply()
    If this.Applying Then Exit Sub               '<~ could break accidental recursion... leave this here...
    
    If this.Mode = OneTimeBinding Then
        If Not this.Applied Then ApplyToTarget
        
    ElseIf this.Mode = OneWayBinding Or this.Mode = TwoWayBinding Then
        ApplyToTarget
        
    ElseIf this.Mode = OneWayToSource Then
        ApplyToSource
        
    End If

End Sub

'@Description "Reads from the source and writes to the target."
Public Sub ApplyToTarget()
    
    If Not CanApplyToTarget Then Exit Sub
    this.Applying = True
    
    Dim SourceValue As Variant
    Dim UseFallbackDefault As Boolean
    
    If this.Source.Object Is Nothing Then
        UseFallbackDefault = TryGetDefaultBindingValue(outValue:=SourceValue)
    End If
    
    this.Source.Resolve
    
    If this.Source.Object Is Nothing And Not UseFallbackDefault Then
        DebugMessage "ApplyToTarget aborted: source object in path '" & this.Source.Path & _
                     "' is Nothing and there is no fallback default value for this binding."
        this.Applying = False
        Exit Sub
        
    ElseIf this.Source.Object Is Nothing Then
        DebugMessage "ApplyToTarget: source object in path '" & this.Source.Path & _
                     "' is Nothing; binding target property '" & this.Target.PropertyName & "' to default/fallback value."
        
    Else
        If Not this.Source.TryReadPropertyValue(outValue:=SourceValue) Then
            DebugMessage "ApplyToTarget failed to read current source property value."
            this.Applying = False
            Exit Sub
        End If
    End If
    
    this.Target.Resolve
    
    Dim CurrentValue As Variant: CurrentValue = Empty
    If Not this.Target.TryReadPropertyValue(outValue:=CurrentValue) Then
        DebugMessage "ApplyToTarget failed to read current target property value."
        this.Applying = False
        Exit Sub
    End If
    
    Dim Proceed As Boolean
    If Not IsArray(SourceValue) And Not IsArray(CurrentValue) Then
        Proceed = SourceValue <> CurrentValue
    Else
        Proceed = True
    End If
    
    If Proceed Then
        
        If Not this.StringFormat Is Nothing Then
            SourceValue = this.StringFormat.Format(SourceValue)
        End If
        
        If Not this.Target.TryWritePropertyValue(SourceValue) Then
            DebugMessage "ApplyToTarget failed to write to target property."
        Else
            DebugMessage "ApplyToTarget successfully applied to target."
            this.Applied = True
        End If
    End If
    this.Applying = False
    
End Sub

'@Description "Reads from the target and writes to the source."
Public Function ApplyToSource() As ApplyResult

    On Error GoTo CleanFail

    If Not CanApplyToSource Then Exit Function
    this.Applying = True
    
    this.Target.Resolve
    Dim TargetValue As Variant
    If Not this.Target.TryReadPropertyValue(outValue:=TargetValue) Then
        DebugMessage "ApplyToSource failed to read target property value."
        ApplyToSource = ApplyResult.BindingFailed
        this.Applying = False
        Exit Function
    End If
    
    this.Source.Resolve
    
    Dim CurrentValue As Variant
    If Not this.Source.TryReadPropertyValue(outValue:=CurrentValue) Then
        DebugMessage "ApplyToSource failed to read current source property value."
        ApplyToSource = ApplyResult.BindingFailed
        this.Applying = False
        Exit Function
    End If
    
    Dim Proceed As Boolean
    If Not IsArray(CurrentValue) Then
        Proceed = TargetValue <> CurrentValue
    Else
        Proceed = True
    End If
    
    If Proceed Then
        
        If Not this.Source.TryWritePropertyValue(TargetValue) Then
            DebugMessage "ApplyToSource failed to write to source property."
            ApplyToSource = ApplyResult.BindingFailed
            Exit Function
        End If
        
        DebugMessage "Binding was successfully applied to source."
        ApplyToSource = ApplyResult.BindingSuccess
        
    End If
    
CleanExit:
    this.Applying = False
    Exit Function
    
CleanFail:
    DebugMessage "ApplyToSource failed unexpectedly."
    ApplyToSource = ApplyResult.BindingFailed
    Resume CleanExit
    Resume
    
End Function

'@Description "Gets a default value for certain specific target properties, used when source path cannot be fully resolved"
'e.g. when target binds to "SomeObjectProperty.SomeProperty" and "SomeObjectProperty" is Nothing.
Private Function TryGetDefaultBindingValue(ByRef outValue As Variant) As Boolean

    Select Case this.Target.PropertyName
    
    Case "Text", "Caption"
        outValue = VBA.Constants.vbNullString
        TryGetDefaultBindingValue = True
            
    Case "Enabled", "Visible"
        outValue = False
        TryGetDefaultBindingValue = True
            
    Case "Value"
        If TypeOf this.Target Is MSForms.CheckBox Or TypeOf this.Target Is MSForms.OptionButton Then
            outValue = False
            TryGetDefaultBindingValue = True
        End If
            
    End Select
    
End Function

Private Sub Dispose()

    If this.Disposed Then
        Debug.Print VBA.Information.TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If
    
    If Not this.Context Is Nothing Then
        Set this.Context = Nothing
    End If
    
    If Not this.Handlers Is Nothing Then
        Set this.Handlers = Nothing
    End If
    
    this.Disposed = True
    
End Sub

Private Sub Class_Terminate()
    If Not this.Disposed Then Dispose
End Sub


Private Sub IDisposable_Dispose()
    Dispose
End Sub

Private Sub IPropertyBinding_Apply()
    Apply
End Sub

Private Property Get IPropertyBinding_DefaultTargetProperty() As String
    IPropertyBinding_DefaultTargetProperty = DefaultTargetProperty
End Property

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = this.Mode
End Property

Private Property Get IPropertyBinding_Source() As IBindingPath
    Set IPropertyBinding_Source = this.Source
End Property

Private Property Get IPropertyBinding_StringFormat() As IStringFormatter
    Set IPropertyBinding_StringFormat = this.StringFormat
End Property

Private Property Get IPropertyBinding_Target() As IBindingPath
    Set IPropertyBinding_Target = this.Target
End Property

Private Property Get IPropertyBinding_UpdateSourceTrigger() As BindingUpdateSourceTrigger
    IPropertyBinding_UpdateSourceTrigger = this.UpdateSourceTrigger
End Property


