VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Guard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Class for common guards."
'@IgnoreModule ProcedureNotUsed
'@Folder "Guard"
'@PredeclaredId
'@ModuleDescription("Class for common guards.")
Option Explicit
Option Compare Text

'@Ignore ProcedureNotUsed
'@Description("Returns instance reference")
Public Property Get Self() As Guard
Attribute Self.VB_Description = "Returns instance reference"
    Set Self = Me
End Property


'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As Guard
Attribute Class.VB_Description = "Returns class reference"
    Set Class = Guard
End Property


'@Description("Default class factory")
Public Function Create() As Guard
Attribute Create.VB_Description = "Default class factory"
    Dim result As Guard
    Set result = Guard 'For singleton, return the predeclared class object
    Set Create = result
End Function


Private Sub Class_Initialize()
    Dim this As TError
    With this
        .trapped = Not Me Is Guard.Self 'Enforce the singleton pattern
        If .trapped Then
            .Name = "SingletonErr"
            .Number = ErrNo.SingletonErr
            .source = TypeName(Me) & " type"
            .message = "Singleton already instantiated"
            .Description = "Creation of non-default instances is not supported. Use predeclared instance."
            RaiseError this
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified string is empty.")
Public Sub EmptyString(ByVal Text As Variant)
Attribute EmptyString.VB_Description = "Raises a run-time error if the specified string is empty."
    Dim errorDetails As TError
    With errorDetails
        .trapped = (VarType(Text) <> vbString)
        If .trapped Then
            .Name = "TypeMismatchErr"
            .Number = ErrNo.TypeMismatchErr
            .source = TypeName(Text) & " type"
            .message = "String required"
            .Description = "Variable type: " & TypeName(Text) & ". String required."
            RaiseError errorDetails
        End If
        
        .trapped = (Text = vbNullString)
        If .trapped Then
            .Name = "EmptyStringErr"
            .Number = ErrNo.EmptyStringErr
            .source = "String variable"
            .message = "String variable empty"
            .Description = .message
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified object reference is Nothing.")
Public Sub NullReference(ByVal instanceVar As Variant)
Attribute NullReference.VB_Description = "Raises a run-time error if the specified object reference is Nothing."
    Dim errorDetails As TError
    With errorDetails
        .trapped = Not IsObject(instanceVar)
        If .trapped Then
            .Name = "ObjectRequiredErr"
            .Number = ErrNo.ObjectRequiredErr
            .source = TypeName(instanceVar) & " Type"
            .message = "Object required"
            .Description = "Variable type: " & TypeName(instanceVar) & ". Object required."
            RaiseError errorDetails
        End If
        
        .trapped = instanceVar Is Nothing
        If .trapped Then
            .Name = "ObjectNotSetErr"
            .Number = ErrNo.ObjectNotSetErr
            .source = "Variable is Nothing"
            .message = "Object variable not set"
            .Description = "Variable type: " & TypeName(instanceVar) & ". Object not set."
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified object reference is not Nothing.")
Public Sub NonNullReference(ByVal instanceVar As Variant)
Attribute NonNullReference.VB_Description = "Raises a run-time error if the specified object reference is not Nothing."
    Dim errorDetails As TError
    With errorDetails
        .trapped = Not IsObject(instanceVar)
        If .trapped Then
            .Name = "ObjectRequiredErr"
            .Number = ErrNo.ObjectRequiredErr
            .source = TypeName(instanceVar) & " Type"
            .message = "Object required"
            .Description = "Variable type: " & TypeName(instanceVar) & ". Object required."
            RaiseError errorDetails
        End If
        
        .trapped = Not instanceVar Is Nothing
        If .trapped Then
            .Name = "ObjectSetErr"
            .Number = ErrNo.ObjectSetErr
            .source = TypeName(instanceVar) & " Type"
            .message = "Object variable already set"
            .Description = "Variable type: " & TypeName(instanceVar) & ". Object is not Nothing."
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified instance isn't the default instance (singleton).")
Public Sub Singleton(ByVal instanceVar As Object)
Attribute Singleton.VB_Description = "Raises a run-time error if the specified instance isn't the default instance (singleton)."
    Debug.Assert IsObject(instanceVar)
    Guard.NullReference instanceVar
        
    Dim classVar As Object
    Set classVar = instanceVar.Class
    
    Dim errorDetails As TError
    With errorDetails
        .trapped = Not instanceVar Is classVar
        If .trapped Then
            .Name = "SingletonErr"
            .Number = ErrNo.SingletonErr
            .source = TypeName(instanceVar) & " type"
            .message = "Singleton already instantiated"
            .Description = "Creation of non-default instances is not supported. Use predeclared instance."
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified instance isn't the default instance.")
Public Sub NonDefaultInstance(ByVal instanceVar As Object)
Attribute NonDefaultInstance.VB_Description = "Raises a run-time error if the specified instance isn't the default instance."
    Guard.NullReference instanceVar
    
    Dim classVar As Object
    Set classVar = instanceVar.Class
    
    Dim errorDetails As TError
    With errorDetails
        .trapped = Not instanceVar Is classVar
        If .trapped Then
            .Name = "NonDefaultInstanceErr"
            .Number = ErrNo.NonDefaultInstanceErr
            .source = TypeName(classVar) & " Class"
            .message = "Default (" & TypeName(classVar) & ") object instance must be used"
            .Description = vbNullString
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified expression is false.")
Public Sub Expression(ByVal AssertionResult As Boolean, ByVal source As String, ByVal message As String)
Attribute Expression.VB_Description = "Raises a run-time error if the specified expression is false."
    Dim errorDetails As TError
    With errorDetails
        .trapped = Not AssertionResult
        If .trapped Then
            .Name = "CustomErr"
            .Number = ErrNo.CustomErr
            .source = source
            .message = message
            .Description = vbNullString
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified instance is the default instance.")
Public Sub DefaultInstance(ByVal instanceVar As Object)
Attribute DefaultInstance.VB_Description = "Raises a run-time error if the specified instance is the default instance."
    Guard.NullReference instanceVar
    
    Dim classVar As Object
    Set classVar = instanceVar.Class
    
    Dim errorDetails As TError
    With errorDetails
        .trapped = instanceVar Is classVar
        If .trapped Then
            .Name = "DefaultInstanceErr"
            .Number = ErrNo.DefaultInstanceErr
            .source = TypeName(classVar) & " Class"
            .message = "Non-default (" & TypeName(classVar) & ") object instance must be used"
            .Description = vbNullString
            RaiseError errorDetails
        End If
    End With
End Sub


'@Ignore ProcedureNotUsed
'@Description("Raises a run-time error if the specified file does not exist.")
Public Sub FileNotExist(ByVal FilePathName As String)
Attribute FileNotExist.VB_Description = "Raises a run-time error if the specified file does not exist."
    Dim fsObject As Scripting.FileSystemObject
    Set fsObject = New Scripting.FileSystemObject
        
    Dim errorDetails As TError
    With errorDetails
        .trapped = Not fsObject.FileExists(FilePathName)
        If .trapped Then
            .Name = "FileNotFoundErr"
            .Number = ErrNo.FileNotFoundErr
            .source = FilePathName
            .message = "File not found " & FilePathName
            .Description = vbNullString
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified connection has active transaction. (Trapped on True)")
Public Sub InTransaction(ByVal HasActiveTransaction As Boolean)
Attribute InTransaction.VB_Description = "Raises a run-time error if the specified connection has active transaction. (Trapped on True)"
    Debug.Assert VarType(HasActiveTransaction) = vbBoolean
    
    Dim errorDetails As TError
    With errorDetails
        .trapped = HasActiveTransaction
        If .trapped Then
            .Name = "InTransaction"
            .Number = ErrNo.AdoInTransactionErr
            .source = vbNullString
            .message = "Current database connection already in transaction"
            .Description = "Attempted operation not permmited when transaction in progress (e.g. starting a new transaction)."
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified connection has no active transaction. (Trapped on False)")
Public Sub NotInTransaction(ByVal HasActiveTransaction As Boolean)
Attribute NotInTransaction.VB_Description = "Raises a run-time error if the specified connection has no active transaction. (Trapped on False)"
    Debug.Assert VarType(HasActiveTransaction) = vbBoolean
    
    Dim errorDetails As TError
    With errorDetails
        .trapped = Not HasActiveTransaction
        If .trapped Then
            .Name = "NotInTransaction"
            .Number = ErrNo.AdoNotInTransactionErr
            .source = vbNullString
            .message = "Current database connection is not in transaction"
            .Description = "Attempted operation not permmited when no  active transaction in progress (e.g. commit or rollback)."
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified connection string is not valid or provider cannot be found. (Trapped on Nothing or vbNullString)")
Public Sub ConnectionString(ByVal dboOrConnStr As Variant)
Attribute ConnectionString.VB_Description = "Raises a run-time error if the specified connection string is not valid or provider cannot be found. (Trapped on Nothing or vbNullString)"
    Dim errorDetails As TError
    With errorDetails
        .trapped = IsFalsy(dboOrConnStr)
        If .trapped Then
            .Name = "ConnectionStringError"
            .Number = ErrNo.AdoConnectionStringErr
            .source = "Argument is Nothing"
            .message = "Either the connection string is invalid or provider not found"
            .Description = .message
            RaiseError errorDetails
        End If
    End With
End Sub


'@Description("Raises a run-time error if the specified references are not both arrays having same base and size")
Public Sub IncompatibleArrays(ByVal ArrayA As Variant, ByVal ArrayB As Variant)
Attribute IncompatibleArrays.VB_Description = "Raises a run-time error if the specified references are not both arrays having same base and size"
    Dim errorDetails As TError
    With errorDetails
        .trapped = ((VarType(ArrayA) And VarType(ArrayB) And vbArray) <> vbArray)
        If .trapped Then
            .Name = "Invalid parameters"
            .Number = ErrNo.InvalidParameterErr
            .source = "Bad arrays"
            .message = "Expected two arrays"
            .Description = .message
            RaiseError errorDetails
        End If
    End With
    
    With errorDetails
        .trapped = (LBound(ArrayA) <> LBound(ArrayB))
        If .trapped Then
            .Name = "Incompatible arrays"
            .Number = ErrNo.IncompatibleArraysErr
            .source = "Bad arrays"
            .message = "Provided arrays have different bases"
            .Description = .message
            RaiseError errorDetails
        End If
    End With
    
    With errorDetails
        .trapped = (UBound(ArrayA) <> UBound(ArrayB))
        If .trapped Then
            .Name = "Incompatible arrays"
            .Number = ErrNo.IncompatibleArraysErr
            .source = "Bad arrays"
            .message = "Provided arrays have different sizes"
            .Description = .message
            RaiseError errorDetails
        End If
    End With
End Sub