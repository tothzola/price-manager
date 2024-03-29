VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EscapeSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("System.Logger.Strings")

Private Type tEscapeSequence
    EscapeString As String
    ReplacementString As String
    IsAsciiCharacter As Boolean
    AsciiBase As AsciiEscapeBase
End Type

Public Enum AsciiEscapeBase
    Octal
    Hexadecimal
End Enum

Private this As tEscapeSequence
Option Explicit

Public Property Get EscapeString() As String
    EscapeString = this.EscapeString
End Property

Friend Property Let EscapeString(ByVal Value As String)
    this.EscapeString = Value
End Property

Public Property Get ReplacementString() As String
    ReplacementString = this.ReplacementString
End Property

Friend Property Let ReplacementString(ByVal Value As String)
    this.ReplacementString = Value
End Property

'@Ignore ProcedureNotUsed
Public Property Get IsAsciiCharacter() As Boolean
    IsAsciiCharacter = this.IsAsciiCharacter
End Property

Friend Property Let IsAsciiCharacter(ByVal Value As Boolean)
    this.IsAsciiCharacter = Value
End Property

'@Ignore ProcedureNotUsed
Public Property Get AsciiBase() As AsciiEscapeBase
    AsciiBase = this.AsciiBase
End Property

Friend Property Let AsciiBase(ByVal Value As AsciiEscapeBase)
    this.AsciiBase = Value
End Property

Public Function Create(ByVal Escape As String, ByVal replacement As String, _
                       Optional ByVal isAsciiReplacement As Boolean = False, _
                       Optional ByVal Base As AsciiEscapeBase = Octal) As EscapeSequence

    Dim result As EscapeSequence
    Set result = New EscapeSequence
    
    result.EscapeString = Escape
    result.ReplacementString = replacement
    result.IsAsciiCharacter = isAsciiReplacement
    result.AsciiBase = Base

    Set Create = result

End Function

'@Ignore ProcedureNotUsed
Public Function Execute(ByVal stringValue As String) As String
    
    Dim result As String
    
    If this.IsAsciiCharacter Then
        result = ProcessAsciiEscape(stringValue, this.EscapeString)

    ElseIf StringType.Contains(stringValue, this.EscapeString) Then
        result = Replace(stringValue, this.EscapeString, this.ReplacementString)

    End If

    Execute = result
End Function

Private Function ProcessAsciiEscape(ByVal FormatString As String, _
                                    ByVal regexPattern As String) As String

    If Not StringType.Contains(FormatString, "\") Then Exit Function
    
    Dim Value As String
    
    Dim regex As RegExp
    Set regex = New RegExp
    regex.Pattern = regexPattern
    regex.IgnoreCase = True
    regex.Global = True

    Dim prefix As String
    Select Case this.AsciiBase
    Case AsciiEscapeBase.Octal
        prefix = "&O"

    Case AsciiEscapeBase.Hexadecimal
        prefix = "&H"

    End Select
    
    Dim matches As MatchCollection
    Set matches = regex.Execute(FormatString)
    Dim thisMatch As Match
    For Each thisMatch In matches
        Dim Char As Long
        Char = CLng(prefix & thisMatch.SubMatches.[_Defalut](0))
        Value = Replace(FormatString, thisMatch.Value, VBA.Strings.Chr$(Char))

    Next

    Set regex = Nothing
    Set matches = Nothing

    ProcessAsciiEscape = Value
End Function


