Attribute VB_Name = "InitializeLogger"
'@Folder("Logger")
Option Explicit
Option Private Module
    
Public Const LOGGER_NAME As String = "PriceApprovalManagerLogger"
Private Const LOGFILE As String = "PriceApprovalManagerLogger.log"

'@EntryPoint
Public Sub InitLogger()
    Dim Path As String: Path = ThisWorkbook.Path
    Dim Separator As String: Separator = Excel.Application.PathSeparator
    Dim LoggingFilePath As String: LoggingFilePath = Path & Separator & LOGFILE
    LogManager.Register FileLogger.Create(LOGGER_NAME, DebugLevel, LoggingFilePath)
End Sub

'@EntryPoint
Public Sub LoggerEnabledCheck()
    If Not LogManager.IsEnabled(DebugLevel) Then
        InitLogger
    End If
End Sub
