VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBStringFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object representing a format string in a notation compatible for use with VBA.Strings.Format$()."
'@ModuleDescription "An object representing a format string in a notation compatible for use with VBA.Strings.Format$()."
'@Folder("Structure.StringFormating")
'@PredeclaredId
'@Exposed
Option Explicit
Implements IStringFormatter

Private Type TState
    FormatString As String
End Type

Private this As TState

Public Function Create(ByVal FormatString As String) As IStringFormatter
    Dim result As VBStringFormatter
    Set result = New VBStringFormatter
    result.FormatString = FormatString
    Set Create = result
End Function

Public Property Get FormatString() As String
    FormatString = this.FormatString
End Property

Friend Property Let FormatString(ByVal RHS As String)
    GuardClauses.GuardDefaultInstance Me, VBStringFormatter
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
    IStringFormatter_Format = VBA.Strings.Format$(LocalArgs(LBound(LocalArgs)), this.FormatString)
End Function

Private Property Get IStringFormatter_FormatString() As String
    IStringFormatter_FormatString = this.FormatString
End Property
