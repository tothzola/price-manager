Attribute VB_Name = "StructuresForErrors"
'@Folder "Guard"
Option Explicit

Private Const adErrInvalidParameterType As Long = &HE3D&

Public Enum ErrNo
    PassedNoErr = 0&
    SubscriptOutOfRange = 9&
    TypeMismatchErr = 13&
    FileNotFoundErr = 53&
    ObjectNotSetErr = 91&
    ObjectRequiredErr = 424&
    InvalidObjectUseErr = 425&
    MemberNotExistErr = 438&
    ActionNotSupportedErr = 445&
    InvalidParameterErr = 1004&
    NoObject = 31004&
    
    InvalidFileName = VBA.vbObjectError + 42&
    CustomErr = VBA.vbObjectError + 1000&
    NotImplementedErr = VBA.vbObjectError + 1001&
    IncompatibleArraysErr = VBA.vbObjectError + 1002&
    ObjectAlreadyInitialized = VBA.vbObjectError + 1003&
    DefaultInstanceErr = VBA.vbObjectError + 1011&
    NonDefaultInstanceErr = VBA.vbObjectError + 1012&
    EmptyStringErr = VBA.vbObjectError + 1013&
    SingletonErr = VBA.vbObjectError + 1014&
    UnknownClassErr = VBA.vbObjectError + 1015&
    ObjectSetErr = VBA.vbObjectError + 1091&
    CouldNotOpen = VBA.vbObjectError + 1092&
    UpdateFailed = VBA.vbObjectError + 1093&
    LoggerAlreadyRegistered = VBA.vbObjectError + 1098&
    NoRegisteredLogger = VBA.vbObjectError + 1099&
    Exeption = VBA.vbObjectError + 9001&
    NullException = VBA.vbObjectError + 9002&
    AdoFeatureNotAvailableErr = ADODB.ErrorValueEnum.adErrFeatureNotAvailable
    AdoInTransactionErr = ADODB.ErrorValueEnum.adErrInTransaction
    AdoNotInTransactionErr = ADODB.ErrorValueEnum.adErrInvalidTransaction
    AdoConnectionStringErr = ADODB.ErrorValueEnum.adErrProviderNotFound
    AdoInvalidParameterTypeErr = VBA.vbObjectError + adErrInvalidParameterType
End Enum


Public Type TError
    Number As ErrNo
    Name As String
    Source As String
    Message As String
    Description As String
    trapped As Boolean
End Type


'@Ignore ProcedureNotUsed
'@Description("Re-raises the current error, if there is one.")
Public Sub RethrowOnError()
Attribute RethrowOnError.VB_Description = "Re-raises the current error, if there is one."
    With VBA.Err
        If .Number <> 0 Then
            'Debug.Print "Error " & .Number, .Description
            '.Raise .Number
            MsgBox .Number & " " & .Description, vbCritical, Title:=PriceApprovalSignature
            LogManager.Log ErrorLevel, .Number & vbTab & .Description
        End If
    End With
End Sub


'@Description("Formats and raises a run-time error.")
Public Sub RaiseError(ByRef errorDetails As TError)
Attribute RaiseError.VB_Description = "Formats and raises a run-time error."
    With errorDetails
        Dim Message As Variant
        Message = Array("Error:", _
                        "name: " & .Name, _
                        "number: " & .Number, _
                        "message: " & .Message, _
                        "description: " & .Description, _
                        "source: " & .Source)
        'VBA.Err.Raise .Number, .source, .Message
        MsgBox .Number & " " & .Source & " " & .Message, vbCritical, Title:=PriceApprovalSignature
        LogManager.Log ErrorLevel, Join(Message, vbNewLine & vbTab)
    End With
End Sub


'@Description("Tests if argument is falsy: 0, False, vbNullString, Empty, Null, Nothing")
Public Function IsFalsy(ByVal arg As Variant) As Boolean
Attribute IsFalsy.VB_Description = "Tests if argument is falsy: 0, False, vbNullString, Empty, Null, Nothing"
    Select Case VarType(arg)
        Case vbEmpty, vbNull
            IsFalsy = True
        Case vbInteger, vbLong, vbSingle, vbDouble
            IsFalsy = Not CBool(arg)
        Case vbString
            IsFalsy = (arg = vbNullString)
        Case vbObject
            IsFalsy = (arg Is Nothing)
        Case vbBoolean
            IsFalsy = Not arg
        Case Else
            IsFalsy = False
    End Select
End Function
