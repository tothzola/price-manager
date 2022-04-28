VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContextModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object keeping command and property bindings in scope."
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@ModuleDescription "An object keeping command and property bindings in scope."
'@Folder("Structure")
'@PredeclaredId
'@Exposed
'Copyright (c) 2020 Mathieu Guindon
Option Explicit

Implements IContextModel
Implements IDisposable

Public Enum StringFormatterStrategy
    UseDotNetStringFormatSyntax
    UseVBStringFormatSyntax
End Enum

Private Type TState
    DefaultStringFormatSyntax As StringFormatterStrategy
    
    StringFormatterFactory As IStringFormatterFactory
    Bindings As IBindingPerformer
    
    DebugOutput As Boolean
    Disposed As Boolean
End Type

Private this As TState

'@Description "Creates a new application context."
Public Function Create( _
       Optional ByVal Bindings As IBindingPerformer, _
       Optional ByVal FormatterFactory As IStringFormatterFactory, _
       Optional ByVal DebugOutput As Boolean = False) As IContextModel
Attribute Create.VB_Description = "Creates a new application context."
       
    GuardClauses.GuardNonDefaultInstance Me, ContextModel, VBA.Information.TypeName(Me)

    Dim result As ContextModel
    Set result = New ContextModel
    result.DebugOutput = DebugOutput
    
    'temporal coupling: IStringFormatterVBFactory instance needs to be set before we init the binding manager.
    InitStringFormatterFactory result, FormatterFactory
    
    'more temporal coupling...
    InitBindingPerformer result, Bindings
    
    Set Create = result
End Function

Private Sub InitStringFormatterFactory(ByVal Context As ContextModel, ByVal FormatterFactory As IStringFormatterFactory)

    Dim Factory As IStringFormatterFactory
    If FormatterFactory Is Nothing Then
        Select Case this.DefaultStringFormatSyntax

        Case StringFormatterStrategy.UseDotNetStringFormatSyntax
            Set Factory = New StringFormatterFactory

        Case StringFormatterStrategy.UseVBStringFormatSyntax
            Set Factory = New VBStringFormatterFactory

        End Select
    Else
        Set Factory = FormatterFactory
    End If
    Set Context.StringFormatterFactory = Factory
    
End Sub

Private Sub InitBindingPerformer(ByVal Context As ContextModel, ByVal Bindings As IBindingPerformer)
    Dim Manager As IBindingPerformer
    
    If Bindings Is Nothing Then
    
        Dim ConcreteManager As BindingPerformer
        Set ConcreteManager = BindingPerformer.Create(Context, Context.StringFormatterFactory, DebugOutput:=Context.DebugOutput)
        Set Manager = ConcreteManager

    Else
        'use injected instance (could be a test stub):
        Set Manager = Bindings
    
    End If
    Set Context.Bindings = Manager
End Sub

'@Description "Gets/sets the binding manager reference."
Public Property Get Bindings() As IBindingPerformer
Attribute Bindings.VB_Description = "Gets/sets the binding manager reference."
    GuardClauses.GuardDefaultInstance Me, ContextModel, VBA.Information.TypeName(Me)
    Set Bindings = this.Bindings
End Property

Friend Property Set Bindings(ByVal RHS As IBindingPerformer)
    GuardClauses.GuardDefaultInstance Me, ContextModel, VBA.Information.TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.Bindings, VBA.Information.TypeName(Me)
    GuardClauses.GuardNullReference RHS, VBA.Information.TypeName(Me)
    Set this.Bindings = RHS
End Property

'@Description "Gets/sets the IStringFormatter factory reference."
Public Property Get StringFormatterFactory() As IStringFormatterFactory
Attribute StringFormatterFactory.VB_Description = "Gets/sets the IStringFormatter factory reference."
    GuardClauses.GuardDefaultInstance Me, ContextModel, VBA.Information.TypeName(Me)
    Set StringFormatterFactory = this.StringFormatterFactory
End Property

Friend Property Set StringFormatterFactory(ByVal RHS As IStringFormatterFactory)
    GuardClauses.GuardDefaultInstance Me, ContextModel, VBA.Information.TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.StringFormatterFactory, VBA.Information.TypeName(Me)
    GuardClauses.GuardNullReference RHS, VBA.Information.TypeName(Me)
    Set this.StringFormatterFactory = RHS
End Property

'@Description "Gets/sets a value that determines the StringFormat syntax to use when no implementation is provided."
Public Property Get DefaultStringFormatSyntax() As StringFormatterStrategy
Attribute DefaultStringFormatSyntax.VB_Description = "Gets/sets a value that determines the StringFormat syntax to use when no implementation is provided."
    GuardClauses.GuardNonDefaultInstance Me, ContextModel, VBA.Information.TypeName(Me)
    DefaultStringFormatSyntax = this.DefaultStringFormatSyntax
End Property

Public Property Let DefaultStringFormatSyntax(ByVal RHS As StringFormatterStrategy)
    GuardClauses.GuardNonDefaultInstance Me, ContextModel, VBA.Information.TypeName(Me)
    this.DefaultStringFormatSyntax = RHS
End Property

'@Description "Gets/sets a value that determines whether binding managers produce debugger output."
Public Property Get DebugOutput() As Boolean
Attribute DebugOutput.VB_Description = "Gets/sets a value that determines whether binding managers produce debugger output."
    DebugOutput = this.DebugOutput
End Property

Public Property Let DebugOutput(ByVal RHS As Boolean)
    this.DebugOutput = RHS
End Property

Private Sub Dispose()
    If this.Disposed Then
        Debug.Print VBA.Information.TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If

    If Not this.StringFormatterFactory Is Nothing Then
        Disposable.TryDispose this.StringFormatterFactory
        Set this.StringFormatterFactory = Nothing
    End If
    
    If Not this.Bindings Is Nothing Then
        Disposable.TryDispose this.Bindings
        Set this.Bindings = Nothing
    End If
    
    this.Disposed = True
    
End Sub

Private Sub Class_Terminate()
    If Not this.Disposed Then Dispose
End Sub

Private Property Get IContextModel_Bindings() As IBindingPerformer
    Set IContextModel_Bindings = this.Bindings
End Property

Private Property Get IContextModel_StringFormatterFactory() As IStringFormatterFactory
    Set IContextModel_StringFormatterFactory = this.StringFormatterFactory
End Property

Private Property Get IContextModel_DebugOutput() As Boolean
    IContextModel_DebugOutput = this.DebugOutput
End Property

Private Sub IDisposable_Dispose()
    Dispose
End Sub
