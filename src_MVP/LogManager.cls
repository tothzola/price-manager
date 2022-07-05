VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LogManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule EmptyElseBlock
'@Folder("System.Logger")
'@PredeclaredId
'@Exposed
Option Explicit

Public Enum LogLevel
    TraceLevel = 0
    DebugLevel
    InfoLevel
    WarnLevel
    ErrorLevel
    FatalLevel
End Enum

Private Type TLogManager
    Formatter As ILogMessageFormatter
    Loggers As New Scripting.Dictionary
    Disposed As Boolean
End Type

Private this As TLogManager

'@Ignore ProcedureNotUsed
Public Property Get Formatter() As ILogMessageFormatter
    Set Formatter = this.Formatter
End Property

'@Ignore ProcedureNotUsed
Public Property Set Formatter(ByVal Value As ILogMessageFormatter)
    Set this.Formatter = Value
End Property

Public Sub UnRegister(ByVal Logger As String)
    If this.Loggers.exists(Logger) Then
        this.Loggers.Remove Logger
    Else
        '"There was no logger Initialized"
        Debug.Print "There was no logger Initialized"
    End If
End Sub

Public Sub Register(ByVal Logger As ILogger)
    If Not this.Loggers.exists(Logger.Name) Then
        this.Loggers.Add Logger.Name, Logger
    Else
        Err.Raise LoggerAlreadyRegistered, "LogManager.Register", "There is already a logger registered with name '" & Logger.Name & "'."
    End If
End Sub

Public Function IsEnabled(ByVal level As LogLevel) As Boolean

    Dim Logger As ILogger
    Dim Item As Variant
    For Each Item In this.Loggers.Items
        Set Logger = Item
        If level >= Logger.MinLevel Then
            IsEnabled = True
            Exit Function
        End If
    Next

End Function

Public Sub Log(ByVal level As LogLevel, ByVal message As String, Optional ByVal loggerName As String)

    Dim Logger As ILogger
    If loggerName = VBA.Constants.vbNullString Then

        Dim Item As Variant
        For Each Item In this.Loggers.Items
            Set Logger = Item
            LogWith Logger, level, message
        Next

    ElseIf this.Loggers.exists(loggerName) Then
        LogWith this.Loggers.Item(loggerName), level, message
    Else
        Err.Raise NoRegisteredLogger, "LogManager.Log", "There is no registered logger named '" & loggerName & "'."
    End If

End Sub

Private Sub LogWith(ByVal Logger As ILogger, ByVal level As LogLevel, ByVal message As String)
    If level >= Logger.MinLevel Then
        Logger.Log FormatMessage(level, Logger.Name, message)
    End If
End Sub

Friend Function FormatMessage(ByVal level As LogLevel, ByVal loggerName As String, ByVal message As String) As String
    FormatMessage = this.Formatter.FormatMessage(level, loggerName, message)
End Function

Private Sub Class_Initialize()
    Set this.Formatter = New DefaultLogMessageFormatter
End Sub

