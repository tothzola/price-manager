VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("System.Logger")
'@PredeclaredId
'@Exposed

Option Explicit

Private Type TFileLogger
    Name As String
    MinLevel As LogLevel
    LogFile As String
    Writer As TextWriter
End Type

Private this As TFileLogger

Implements ILogger

Public Function Create(ByVal LoggerName As String, ByVal loggerMinLevel As LogLevel, ByVal Path As String) As ILogger

    Dim result As FileLogger
    Set result = New FileLogger
    result.Name = LoggerName
    result.MinLevel = loggerMinLevel
    Set result.Writer = TextWriter.Create
    result.LogFile = Path
    Set Create = result

End Function

'@Ignore ProcedureNotUsed
Friend Property Get Name() As String
    Name = this.Name
End Property

Friend Property Let Name(ByVal Value As String)
    this.Name = Value
End Property

'@Ignore ProcedureNotUsed
Friend Property Get MinLevel() As LogLevel
    MinLevel = this.MinLevel
End Property

Friend Property Let MinLevel(ByVal Value As LogLevel)
    this.MinLevel = Value
End Property

'@Ignore ProcedureNotUsed
Friend Property Get LogFile() As String
    LogFile = this.LogFile
End Property

Friend Property Let LogFile(ByVal Value As String)
    this.LogFile = Value
End Property

'@Ignore ProcedureNotUsed
Friend Property Get Writer() As TextWriter
    Set Writer = this.Writer
End Property

Friend Property Set Writer(ByVal Value As TextWriter)
    Set this.Writer = Value
End Property

Private Sub ILogger_Log(ByVal output As String)
    If this.Writer.OpenFile(this.LogFile) Then

        this.Writer.WriteLine output
        this.Writer.CloseFile

    Else
        Err.Raise CouldNotOpen, "ILogger.Log", "FileLogger.LogFile could not be opened."
    End If
End Sub

Private Property Get ILogger_MinLevel() As LogLevel
    ILogger_MinLevel = this.MinLevel
End Property

Private Property Get ILogger_Name() As String
    ILogger_Name = this.Name
End Property

