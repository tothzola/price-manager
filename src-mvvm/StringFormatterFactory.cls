VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringFormatterFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "A factory that creates new .NET-syntax string formatters."
'@ModuleDescription "A factory that creates new .NET-syntax string formatters."
'@Folder("Structure.StringFormating")
'Copyright (c) 2020 Mathieu Guindon
Option Explicit
Implements IStringFormatterFactory

Private Function IStringFormatterFactory_Create(ByVal FormatString As String) As IStringFormatter
    Set IStringFormatterFactory_Create = StringFormatter.Create(FormatString)
End Function

