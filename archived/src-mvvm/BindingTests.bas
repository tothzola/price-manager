Attribute VB_Name = "BindingTests"
'@Folder("Tests.Bindings")
'@TestModule
Option Explicit
Option Private Module

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.AssertClass
#End If

Private Type TState
    ExpectedErrNumber As Long
    ExpectedErrSource As String
    ExpectedErrorCaught As Boolean
    
    Context As IContextModel
    
    ConcreteSUT As BindingPerformer
    AbstractSUT As IBindingPerformer

    BindingSource As TestBindingObject
    BindingTarget As TestBindingObject
    
    SourceProperty As String
    TargetProperty As String
    
    SourcePropertyPath As String
    TargetPropertyPath As String
    
End Type

Private Test As TState

'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        'requires HKCU registration of the Rubberduck COM library.
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        'requires project reference to the Rubberduck COM library.
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    
    Set Test.Context = ContextModel.Create(DebugOutput:=True)
    
    Set Test.ConcreteSUT = Test.Context.Bindings

    Set Test.AbstractSUT = Test.ConcreteSUT

    Set Test.BindingSource = New TestBindingObject
    Set Test.BindingTarget = New TestBindingObject

    Test.SourcePropertyPath = "TestStringProperty"
    Test.TargetPropertyPath = "TestStringProperty"
    Test.SourceProperty = "TestStringProperty"
    Test.TargetProperty = "TestStringProperty"
    
End Sub

'@TestCleanup
Private Sub TestCleanup()

    Set Test.Context = Nothing
    
    Set Test.ConcreteSUT = Nothing
    Set Test.AbstractSUT = Nothing

    Set Test.BindingSource = Nothing
    Set Test.BindingTarget = Nothing
    
    Test.SourcePropertyPath = vbNullString
    Test.TargetPropertyPath = vbNullString
    Test.ExpectedErrNumber = 0
    Test.ExpectedErrorCaught = False
    Test.ExpectedErrSource = vbNullString
    
End Sub

Private Sub ExpectError()
    Dim Message As String
    If Err.Number = Test.ExpectedErrNumber Then
        If (Test.ExpectedErrSource = vbNullString) Or (Err.Source = Test.ExpectedErrSource) Then
            Test.ExpectedErrorCaught = True
        Else
            Message = "An error was raised, but not from the expected source. " & _
                      "Expected: '" & TypeName(Test.ConcreteSUT) & "'; Actual: '" & Err.Source & "'."
        End If
    ElseIf Err.Number <> 0 Then
        Message = "An error was raised, but not with the expected number. Expected: '" & Test.ExpectedErrNumber & "'; Actual: '" & Err.Number & "'."
    Else
        Message = "No error was raised."
    End If
    
    If Not Test.ExpectedErrorCaught Then Assert.Fail Message
End Sub

'@TestMethod("GuardClauses")
Private Sub Create_GuardsNonDefaultInstance()
    Test.ExpectedErrNumber = GuardClauseErrors.InvalidFromNonDefaultInstance
    With New BindingPerformer
        On Error Resume Next
        '@Ignore FunctionReturnValueDiscarded, FunctionReturnValueNotUsed
        .Create Test.Context, New StringFormatterFactory
        ExpectError
        On Error GoTo 0
    End With
End Sub

Private Function DefaultPropertyPathBindingFor(ByVal ProgID As String, ByRef outTarget As Object) As IPropertyBinding
    Set outTarget = CreateObject(ProgID)
    Set DefaultPropertyPathBindingFor = Test.AbstractSUT.BindPropertyPath(Test.BindingSource, _
                                                                          Test.SourcePropertyPath, _
                                                                          outTarget, _
                                                                          Test.TargetPropertyPath)
End Function

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_FrameTargetCreatesOneWayBindingWithNonDefaultTarget()
    Test.TargetPropertyPath = "Font.Bold"
    Dim Target As Object
    Dim result As IPropertyBinding
    Set result = DefaultPropertyPathBindingFor("Forms.Frame.1", outTarget:=Target)
    Assert.AreEqual VBA.TypeName(OneWayPropertyBinding), VBA.TypeName(result), "Actual: " & VBA.TypeName(result)
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_LabelTargetCreatesOneWayBindingWithNonDefaultTarget()
    Test.TargetPropertyPath = "Font.Bold"
    Dim Target As Object
    Dim result As IPropertyBinding
    Set result = DefaultPropertyPathBindingFor("Forms.Label.1", outTarget:=Target)
    Assert.IsTrue TypeOf result Is OneWayPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_FrameTargetBindsCaptionPropertyByDefault()
    Test.TargetPropertyPath = vbNullString
    Dim Target As Object
    Dim result As IPropertyBinding
    Set result = DefaultPropertyPathBindingFor("Forms.Frame.1", outTarget:=Target)
    Assert.IsTrue TypeOf result Is CaptionPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_LabelTargetBindsCaptionPropertyByDefault()
    Test.TargetPropertyPath = vbNullString
    Dim Target As Object
    Dim result As IPropertyBinding
    Set result = DefaultPropertyPathBindingFor("Forms.Label.1", outTarget:=Target)
    Assert.IsTrue TypeOf result Is CaptionPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_NonControlTargetCreatesOneWayBinding()
    Dim result As IPropertyBinding
    Set result = Test.AbstractSUT.BindPropertyPath(Test.BindingSource, Test.SourcePropertyPath, Test.BindingTarget, Test.TargetPropertyPath)
    Assert.IsTrue TypeOf result Is OneWayPropertyBinding
End Sub

'@TestMethod("DefaultPropertyPathBindings")
Private Sub BindPropertyPath_NonControlTargetRequiresTargetPropertyPath()
    Test.ExpectedErrNumber = GuardClauseErrors.StringCannotBeEmpty
    On Error Resume Next
    Test.AbstractSUT.BindPropertyPath _
        Test.BindingSource, _
        Test.SourcePropertyPath, _
        Test.BindingTarget, _
        TargetProperty:=vbNullString
    ExpectError
    On Error GoTo 0
End Sub

'@TestMethod("CallbackPropagation")
Private Sub BindPropertyPath_AddsToPropertyBindingsCollection()
    Dim result As IPropertyBinding
    Set result = Test.AbstractSUT.BindPropertyPath(Test.BindingSource, Test.SourcePropertyPath, Test.BindingTarget, Test.TargetPropertyPath)
    Assert.AreEqual 1, Test.ConcreteSUT.PropertyBindings.Count
    Assert.AreSame result, Test.ConcreteSUT.PropertyBindings.Item(1)
End Sub
