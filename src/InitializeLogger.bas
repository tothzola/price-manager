Attribute VB_Name = "InitializeLogger"
'@Folder("System.Logger")
Option Explicit
Option Private Module
    
Public Const LOGGER_NAME As String = "PriceApprovalLogger_MVP"
Private Const LogFile As String = "PriceApproval_MVP.log"

'@EntryPoint
Public Sub InitLogger()
    Dim Path As String: Path = ThisWorkbook.Path
    Dim Separator As String: Separator = Excel.Application.PathSeparator
    Dim LoggingFilePath As String: LoggingFilePath = Path & Separator & LogFile
    LogManager.Register FileLogger.Create(LOGGER_NAME, DebugLevel, LoggingFilePath)
End Sub

'@EntryPoint
Public Sub LoggerEnabledCheck()
    If Not LogManager.IsEnabled(DebugLevel) Then
        InitLogger
    End If
End Sub
