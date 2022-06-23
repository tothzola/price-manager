Attribute VB_Name = "GuardClauses"
Attribute VB_Description = "Global procedures for throwing custom run-time errors in guard clauses."
'@IgnoreModule ProcedureNotUsed
'@ModuleDescription("Global procedures for throwing custom run-time errors in guard clauses.")
'@Folder("CustomErrors")
Option Explicit
Option Private Module

Private Const CustomErr = VBA.vbObjectError + 1000&

Public Enum GuardClauseErrors
    InvalidFromNonDefaultInstance = CustomErr
    InvalidFromDefaultInstance
    ObjectAlreadyInitialized
    ObjectCannotBeNothing
    StringCannotBeEmpty
End Enum

'@Description("Raises a run-time error if the specified Boolean expression is True.")
Public Sub GuardExpression(ByVal Throw As Boolean, _
                           Optional ByVal Source As String = VBA.Constants.vbNullString, _
                           Optional ByVal Message As String = "Invalid procedure call or argument.", _
                           Optional ByVal ErrNumber As Long = CustomErr)
Attribute GuardExpression.VB_Description = "Raises a run-time error if the specified Boolean expression is True."
    If Throw Then VBA.Information.Err.Raise ErrNumber, Source, Message
End Sub

'@Description("Raises a run-time error if the specified instance isn't the default instance.")
Public Sub GuardNonDefaultInstance(ByVal Instance As Object, ByVal DefaultInstance As Object, _
                                   Optional ByVal Source As String = VBA.Constants.vbNullString, _
                                   Optional ByVal Message As String = "Method should be invoked from the default/predeclared instance of this class.")
Attribute GuardNonDefaultInstance.VB_Description = "Raises a run-time error if the specified instance isn't the default instance."
    Debug.Assert VBA.Information.TypeName(Instance) = VBA.Information.TypeName(DefaultInstance)
    GuardExpression Not Instance Is DefaultInstance, IIf(Source = VBA.Constants.vbNullString, VBA.Information.TypeName(Instance), Source), Message, InvalidFromNonDefaultInstance
End Sub

'@Description("Raises a run-time error if the specified instance is the default instance.")
Public Sub GuardDefaultInstance(ByVal Instance As Object, ByVal DefaultInstance As Object, _
                                Optional ByVal Source As String = VBA.Constants.vbNullString, _
                                Optional ByVal Message As String = "Method should be invoked from a new instance of this class.")
Attribute GuardDefaultInstance.VB_Description = "Raises a run-time error if the specified instance is the default instance."
    Debug.Assert VBA.Information.TypeName(Instance) = VBA.Information.TypeName(DefaultInstance)
    GuardExpression Instance Is DefaultInstance, Source, Message, InvalidFromDefaultInstance
End Sub

'@Description("Raises a run-time error if the specified object reference is already set.")
Public Sub GuardDoubleInitialization(ByVal Value As Variant, _
                                     Optional ByVal Source As String = VBA.Constants.vbNullString, _
                                     Optional ByVal Message As String = "Value is already initialized.")
Attribute GuardDoubleInitialization.VB_Description = "Raises a run-time error if the specified object reference is already set."
    Dim Throw As Boolean
    If IsObject(Value) Then
        Throw = Not Value Is Nothing
    Else
        Throw = Value <> GetDefaultValue(VarType(Value))
    End If
    GuardExpression Throw, Source, Message, ObjectAlreadyInitialized
End Sub

Private Function GetDefaultValue(ByVal VType As VbVarType) As Variant
    Select Case VType
    Case VbVarType.vbString
        GetDefaultValue = VBA.Constants.vbNullString
    Case VbVarType.vbBoolean
        GetDefaultValue = False
    Case VbVarType.vbByte, VbVarType.vbCurrency, VbVarType.vbDate, _
         VbVarType.vbDecimal, VbVarType.vbDouble, VbVarType.vbLong, _
         VbVarType.vbLong, VbVarType.vbSingle
        GetDefaultValue = 0
    Case VbVarType.vbArray, VbVarType.vbEmpty, VbVarType.vbVariant
        GetDefaultValue = Empty
    Case VbVarType.vbNull
        GetDefaultValue = Null
    Case VbVarType.vbObject
        Set GetDefaultValue = Nothing
        #If VBA7 Then
            #If Win64 Then
            Case VbVarType.vbLongLong
                GetDefaultValue = 0
            #Else
            Case VbVarType.vbLong
                GetDefaultValue = 0
            #End If
        #End If
    End Select
End Function

'@Description("Raises a run-time error if the specified object reference is Nothing.")
Public Sub GuardNullReference(ByVal Instance As Object, _
                              Optional ByVal Source As String = VBA.Constants.vbNullString, _
                              Optional ByVal Message As String = "Object reference cannot be Nothing.")
Attribute GuardNullReference.VB_Description = "Raises a run-time error if the specified object reference is Nothing."
    GuardExpression Instance Is Nothing, Source, Message, ObjectCannotBeNothing
End Sub

'@Description("Raises a run-time error if the specified string is empty.")
Public Sub GuardEmptyString(ByVal Value As String, _
                            Optional ByVal Source As String = VBA.Constants.vbNullString, _
                            Optional ByVal Message As String = "String cannot be empty.")
Attribute GuardEmptyString.VB_Description = "Raises a run-time error if the specified string is empty."
    GuardExpression Value = VBA.Constants.vbNullString, Source, Message, StringCannotBeEmpty
End Sub
