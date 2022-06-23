VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestBindingObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An object that can stub a binding source or target for unit tests."
'@ModuleDescription "An object that can stub a binding source or target for unit tests."
'@Folder("Tests.Stubs")
'@PredeclaredId
Option Explicit

Implements ITestStub

Private Type TState
    Stub As TestStubBase

    TestStringProperty As String
    TestNumericProperty As Long
    
    TestBindingObjectProperty As TestBindingObject
End Type

Private this As TState

Public Function Create() As TestBindingObject
    GuardClauses.GuardNonDefaultInstance Me, TestBindingObject, TypeName(Me)
    Dim result As TestBindingObject
    Set result = New TestBindingObject
    Set Create = result
End Function

Public Property Get TestStringProperty() As String
    this.Stub.OnInvoke "TestStringProperty.Get"
    TestStringProperty = this.TestStringProperty
End Property

Public Property Let TestStringProperty(ByVal RHS As String)
    this.Stub.OnInvoke "TestStringProperty.Let"
    If this.TestStringProperty <> RHS Then
        this.TestStringProperty = RHS

    End If
End Property

Public Property Get TestNumericProperty() As Long
    this.Stub.OnInvoke "TestNumericProperty.Get"
    TestNumericProperty = this.TestNumericProperty
End Property

Public Property Let TestNumericProperty(ByVal RHS As Long)
    this.Stub.OnInvoke "TestNumericProperty.Let"
    If this.TestNumericProperty <> RHS Then
        this.TestNumericProperty = RHS
    
    End If
End Property

Public Property Get TestBindingObjectProperty() As TestBindingObject
    this.Stub.OnInvoke "TestBindingObjectProperty.Get"
    Set TestBindingObjectProperty = this.TestBindingObjectProperty
End Property

Public Property Set TestBindingObjectProperty(ByVal RHS As TestBindingObject)
    this.Stub.OnInvoke "TestBindingObjectProperty.Set"
    If Not this.TestBindingObjectProperty Is RHS Then
        Set this.TestBindingObjectProperty = RHS

    End If
End Property

Private Sub Class_Initialize()
    Set this.Stub = New TestStubBase

End Sub

Private Property Get ITestStub_MemberInvokes(ByVal MemberName As String) As Long
    ITestStub_MemberInvokes = this.Stub.MemberInvokes(MemberName)
End Property

Private Function ITestStub_ToString() As String
    ITestStub_ToString = this.Stub.ToString
End Function

Private Sub ITestStub_Verify(ByVal Assert As Object, ByVal MemberName As String, ByVal ExpectedInvokes As Long)
    this.Stub.Verify Assert, MemberName, ExpectedInvokes
End Sub
