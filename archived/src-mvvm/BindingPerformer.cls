VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BindingPerformer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object responsible for managing property bindings."
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@ModuleDescription "An object responsible for managing property bindings."
'@Folder("Structure.Bindings")
'@PredeclaredId
'@Exposed
'Copyright (c) 2020 Mathieu Guindon
Option Explicit

Implements IBindingPerformer
Implements IDisposable

Private Type TState
    Context As IContextModel
    StringFormatterFactory As IStringFormatterFactory
    BindingTargetStrategies As Object
    PropertyBindings As VBA.Collection
    DebugOutput As Boolean
    Disposed As Boolean
End Type

Private this As TState

Private Sub DebugMessage(ByVal Message As String)
    If this.DebugOutput Then
        Debug.Print Message
    End If
End Sub

Public Function Create(ByVal Context As IContextModel, ByVal StringFormatFactory As IStringFormatterFactory, _
                       Optional ByVal DebugOutput As Boolean = False) As IBindingPerformer
                       
    GuardClauses.GuardNonDefaultInstance Me, BindingPerformer, VBA.Information.TypeName(Me)
    
    With New BindingPerformer

        If Context Is Nothing Then
            Set .Context = ContextModel.Create(DebugOutput:=DebugOutput)
        Else
            Set .Context = Context
        End If
        
        .DebugOutput = Context.DebugOutput
        
        If StringFormatFactory Is Nothing Then
            Set .StringFormatterFactory = New StringFormatterFactory
        Else
            Set .StringFormatterFactory = StringFormatFactory
        End If
        
        Set Create = .Self

    End With
    
End Function

Public Property Get Self() As IBindingPerformer
    Set Self = Me
End Property

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is BindingPerformer
End Property

Private Sub Apply(ByVal Source As Object)
    GuardClauses.GuardDefaultInstance Me, BindingPerformer, VBA.Information.TypeName(Me), "Member call is invalid against stateless default instance."
    Dim Binding As IPropertyBinding
    For Each Binding In this.PropertyBindings
        If Source Is Binding.Source.Context Then Binding.Apply
    Next
    
End Sub

Public Property Get Context() As IContextModel
    Set Context = this.Context
End Property

Friend Property Set Context(ByVal RHS As IContextModel)
    GuardClauses.GuardDefaultInstance Me, BindingPerformer, VBA.Information.TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.Context, VBA.Information.TypeName(Me)
    GuardClauses.GuardNullReference RHS, VBA.Information.TypeName(Me)
    Set this.Context = RHS
End Property

Public Property Get PropertyBindings() As VBA.Collection
    Set PropertyBindings = this.PropertyBindings
End Property

Public Property Get StringFormatterFactory() As IStringFormatterFactory
    Set StringFormatterFactory = this.StringFormatterFactory
End Property

Public Property Set StringFormatterFactory(ByVal RHS As IStringFormatterFactory)
    GuardClauses.GuardDefaultInstance Me, BindingPerformer, VBA.Information.TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.StringFormatterFactory, VBA.Information.TypeName(Me)
    GuardClauses.GuardNullReference RHS, VBA.Information.TypeName(Me)
    Set this.StringFormatterFactory = RHS
End Property

Public Property Get DebugOutput() As Boolean
    DebugOutput = this.DebugOutput
End Property

Public Property Let DebugOutput(ByVal RHS As Boolean)
    this.DebugOutput = RHS
End Property

Private Sub Class_Initialize()

    If Not IsDefaultInstance Then
        Set this.BindingTargetStrategies = VBA.CreateObject("Scripting.Dictionary")
        Set this.PropertyBindings = New VBA.Collection
        
        With this.BindingTargetStrategies
            .Add "MSForms.UserForm", New CaptionBindingStrategy
            .Add "MSForms.Frame", New CaptionBindingStrategy
            .Add "MSForms.Label", New CaptionBindingStrategy
            
        End With
    End If
    
End Sub

Private Sub Class_Terminate()
    If Not this.Disposed Then Dispose
End Sub

Private Sub Dispose()

    If this.Disposed Then
        DebugMessage VBA.Information.TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If
    
    If Not this.Context Is Nothing Then
        Set this.Context = Nothing
    End If
    
    If Not this.BindingTargetStrategies Is Nothing Then
        Set this.BindingTargetStrategies = Nothing
    End If

    If Not this.PropertyBindings Is Nothing Then
        '@Ignore VariableNotUsed
        Dim Index As Long
        For Index = 1 To this.PropertyBindings.Count
            Disposable.TryDispose this.PropertyBindings.Item(this.PropertyBindings.Count)
            this.PropertyBindings.Remove this.PropertyBindings.Count
        Next
        Set this.PropertyBindings = Nothing
        
    End If
    this.Disposed = True
    
End Sub

Private Function IBindingPerformer_BindPropertyPath(ByVal Source As Object, ByVal PropertyPath As String, ByVal Target As Object, _
                                                    Optional ByVal TargetProperty As String, _
                                                    Optional ByVal Mode As BindingMode, _
                                                    Optional ByVal UpdateTrigger As BindingUpdateSourceTrigger, _
                                                    Optional ByVal StringFormat As String) As IPropertyBinding
                                                     
    GuardClauses.GuardDefaultInstance Me, BindingPerformer, VBA.Information.TypeName(Me), "Member call is invalid against stateless default instance."

    Dim Formatter As IStringFormatter
    
    If StringFormat <> VBA.Constants.vbNullString And Not this.StringFormatterFactory Is Nothing Then
        Set Formatter = this.StringFormatterFactory.Create(FormatString:=StringFormat)
    End If
    
    Dim Binding As IPropertyBinding
    Dim Strategy As IBindingTargetStrategy
    
    If TryGetBindingStrategyFor(Target:=Target, outStrategy:=Strategy) Then
    
        If TargetProperty = VBA.Constants.vbNullString Then
            
            Set Binding = Strategy.DefaultPropertyBindingFor(Context:=this.Context, _
                                                             Source:=BindingPath.Create(Source, PropertyPath), _
                                                             Target:=Target, _
                                                             Mode:=Mode, _
                                                             UpdateTrigger:=UpdateTrigger, _
                                                             StringFormat:=Formatter)
        Else
            Set Binding = Strategy.PropertyBindingFor(Context:=this.Context, _
                                                      Source:=BindingPath.Create(Source, PropertyPath), _
                                                      Target:=BindingPath.Create(Target, TargetProperty), _
                                                      Mode:=Mode, _
                                                      UpdateTrigger:=UpdateTrigger, _
                                                      StringFormat:=Formatter)
        End If

    Else

        Set Binding = OneWayPropertyBinding.Create(Context:=this.Context, _
                                                   Source:=BindingPath.Create(Source, PropertyPath), _
                                                   Target:=BindingPath.Create(Target, TargetProperty), _
                                                   StringFormat:=Formatter)
                                                 
    End If

    this.PropertyBindings.Add Binding
        
    Set IBindingPerformer_BindPropertyPath = Binding
    
    If Not this.DebugOutput Then Exit Function
    
    On Error Resume Next
    If this.DebugOutput And TargetProperty = VBA.Constants.vbNullString Then
        DebugMessage VBA.Information.TypeName(Me) & ": Binding property path '" & PropertyPath & _
                                                 "' to the default-binding property of type '" & VBA.Information.TypeName(Target) & "'."
    
    ElseIf this.DebugOutput Then
        DebugMessage VBA.Information.TypeName(Me) & ": Binding property path '" & PropertyPath & _
                                                 "' to '" & VBA.Information.TypeName(Target) & "." & TargetProperty & "'."
    
    End If
    On Error GoTo 0
    
End Function

Private Function TryGetBindingStrategyFor(ByVal Target As Object, ByRef outStrategy As IBindingTargetStrategy) As Boolean
    
    Dim LibraryName As String
    
    Dim key As String
    Dim targetLibrary As String
    targetLibrary = VBA.Information.TypeName(Target)
    
    If TypeOf Target Is Excel.Range Then
        key = "Excel" & "." & targetLibrary
        
    ElseIf TypeOf Target Is MSForms.Control Then
        key = "MSForms" & "." & targetLibrary

    ElseIf TypeOf Target Is MSForms.UserForm Then
        key = "MSForms" & "." & "UserForm"
        
    ElseIf TypeOf Target Is Object  And TypeName(Target) = targetLibrary Then
        key = "MSForms" & "." & targetLibrary
        
    Else
        Err.Raise VBA.vbObjectError + 1004&, "BindingPerformer", "Bindingstrategy target type not defined."

    End If
    
    If this.BindingTargetStrategies.exists(key) Then
        Set outStrategy = this.BindingTargetStrategies.Item(key)
        TryGetBindingStrategyFor = True
    End If
    
End Function

Private Sub IBindingPerformer_Apply(ByVal Source As Object)
    GuardClauses.GuardDefaultInstance Me, BindingPerformer, VBA.Information.TypeName(Me)
    Apply Source
End Sub

Private Property Get IBindingPerformer_DebugOutput() As Boolean
    IBindingPerformer_DebugOutput = this.DebugOutput
End Property

Private Property Get IBindingPerformer_StringFormatterFactory() As IStringFormatterFactory
    GuardClauses.GuardDefaultInstance Me, BindingPerformer, VBA.Information.TypeName(Me)
    Set IBindingPerformer_StringFormatterFactory = this.StringFormatterFactory
End Property

Private Sub IDisposable_Dispose()
    Dispose
End Sub

