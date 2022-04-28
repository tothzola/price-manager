VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OneWayPropertyBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A property binding that binds any property of any source, one-way to any property of a UI element target."
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@ModuleDescription "A property binding that binds any property of any source, one-way to any property of a UI element target."
'@Folder("Structure.Bindings.ConcreteBinding")
'@PredeclaredId
'@Exposed
'Copyright (c) 2020 Mathieu Guindon
Option Explicit

Implements IPropertyBinding
Implements IDisposable

Private Type TState
    Base As PropertyBindingBase
    Disposed As Boolean
End Type

Private this As TState

Public Function Create(ByVal Context As IContextModel, ByVal Source As IBindingPath, ByVal Target As IBindingPath, _
                       Optional ByVal StringFormat As IStringFormatter) As IPropertyBinding
    
    Dim BindingBase As PropertyBindingBase
    Set BindingBase = PropertyBindingBase.Create(Context, Source, Target.Context, Target.Path, _
                                                 Mode:=BindingMode.OneWayBinding, _
                                                 UpdateSource:=BindingUpdateSourceTrigger.Never, _
                                                 StringFormat:=StringFormat)
    
    Dim result As OneWayPropertyBinding
    Set result = New OneWayPropertyBinding
    
    result.InjectBindingInfo BindingBase
    
    Set Create = result
    
End Function

Public Sub InjectBindingInfo(ByVal BindingInfo As PropertyBindingBase)
    GuardClauses.GuardDefaultInstance Me, OneWayPropertyBinding, VBA.Information.TypeName(Me)
    GuardClauses.GuardNullReference BindingInfo, VBA.Information.TypeName(Me)
    GuardClauses.GuardDoubleInitialization this.Base, VBA.Information.TypeName(Me)
    
    Set this.Base = BindingInfo
    
End Sub

Private Property Get IsDefaultInstance() As Boolean
    IsDefaultInstance = Me Is OneWayPropertyBinding
End Property


Private Sub IDisposable_Dispose()

    If this.Disposed Then
        Debug.Print VBA.Information.TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If
    
    If Not this.Base Is Nothing Then
        Disposable.TryDispose this.Base
        Set this.Base = Nothing
    End If
    
    this.Disposed = True
    
End Sub

Private Sub IPropertyBinding_Apply()
    this.Base.Apply
End Sub

Private Property Get IPropertyBinding_DefaultTargetProperty() As String
    IPropertyBinding_DefaultTargetProperty = VBA.Constants.vbNullString
End Property

Private Property Get IPropertyBinding_Mode() As BindingMode
    IPropertyBinding_Mode = this.Base.Mode
End Property

Private Property Get IPropertyBinding_Source() As IBindingPath
    Set IPropertyBinding_Source = this.Base.Source
End Property

Private Property Get IPropertyBinding_StringFormat() As IStringFormatter
    Set IPropertyBinding_StringFormat = this.Base.StringFormat
End Property

Private Property Get IPropertyBinding_Target() As IBindingPath
    Set IPropertyBinding_Target = this.Base.Target
End Property

Private Property Get IPropertyBinding_UpdateSourceTrigger() As BindingUpdateSourceTrigger
    IPropertyBinding_UpdateSourceTrigger = this.Base.UpdateSourceTrigger
End Property
