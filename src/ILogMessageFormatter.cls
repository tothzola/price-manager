VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILogMessageFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("System.Logger")
'@Interface
'@Exposed
Option Explicit

Public Function FormatMessage(ByVal level As LogLevel, ByVal LoggerName As String, ByVal Message As String) As String
End Function

