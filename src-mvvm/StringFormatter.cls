VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object representing a format string in a notation compatible for use with .NET String.Format()."
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@ModuleDescription "An object representing a format string in a notation compatible for use with .NET String.Format()."
'@Folder("Structure.StringFormating")
'@PredeclaredId
'@Exposed
'Copyright (c) 2020 Mathieu Guindon
Option Explicit

Implements IStringFormatter

Private Type TState
    FormatString As String
End Type

Private this As TState

Public Function Create(ByVal FormatString As String) As IStringFormatter
    Dim result As StringFormatter
    Set result = New StringFormatter
    result.FormatString = FormatString
    Set Create = result
End Function

Public Property Get FormatString() As String
    FormatString = this.FormatString
End Property

Friend Property Let FormatString(ByVal RHS As String)
    GuardClauses.GuardDefaultInstance Me, StringFormatter
    this.FormatString = RHS
End Property

Private Function IStringFormatter_Format(ParamArray Values() As Variant) As String
    Dim LocalArgs() As Variant

    If IsArray(Values(LBound(Values))) Then
        LocalArgs = Values(LBound(Values))
    Else
        ReDim LocalArgs(LBound(Values) To UBound(Values))
        Dim Index As Long
        For Index = LBound(Values) To UBound(Values)
            LocalArgs(Index) = Values(Index)
        Next
    End If
    With StringBuilder.AppendFormat(this.FormatString, LocalArgs)
        IStringFormatter_Format = .ToString
    End With
End Function

Private Property Get IStringFormatter_FormatString() As String
    IStringFormatter_FormatString = this.FormatString
End Property
