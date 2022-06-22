VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule IndexedUnboundDefaultMemberAccess, ProcedureNotUsed
'@Folder("Logger.Strings")
'@PredeclaredId
'@Exposed
Option Explicit

Private Escapes As VBA.Collection

Private Sub Class_Initialize()
    Dim Factory As EscapeSequence
    Set Factory = New EscapeSequence
    Set Escapes = New VBA.Collection
    Escapes.Add Factory.Create("\n", VBA.Constants.vbNewLine), "0"
    Escapes.Add Factory.Create("\q", VBA.Strings.Chr$(34)), "1"
    Escapes.Add Factory.Create("\t", VBA.Constants.vbTab), "2"
    Escapes.Add Factory.Create("\a", VBA.Strings.Chr$(7)), "3"
    Escapes.Add Factory.Create("\b", VBA.Strings.Chr$(8)), "4"
    Escapes.Add Factory.Create("\v", VBA.Strings.Chr$(13)), "5"
    Escapes.Add Factory.Create("\f", VBA.Strings.Chr$(14)), "6"
    Escapes.Add Factory.Create("\r", VBA.Strings.Chr$(15)), "7"
    Set Factory = Nothing
End Sub

'@Description "VB6 implementation of .net String.Format(), slightly customized."
Public Function Format(ByRef FormatString As String, ParamArray Values()) As String
Attribute Format.VB_Description = "VB6 implementation of .net String.Format(), slightly customized."

    Dim returnValue As String
    Dim valuesCount As Long
        
    'some error-handling constants:
    Const ERR_SOURCE As String = "StringFormat"
    Const ERR_MSG_INVALID_formatString As String = "Invalid format string."
        
    Dim PADDING_CHAR As String
    'use SPACE as default padding character
    If PADDING_CHAR = VBA.Constants.vbNullString Then PADDING_CHAR = VBA.Strings.Chr$(32)
        
    'figure out number of passed values:
    valuesCount = UBound(Values) + 1
                    
    Dim regex As RegExp
    Dim matches As MatchCollection
    Dim thisMatch As Match
    Dim thisString As String
    Dim thisFormat As String
        
    'validate string_format:
    Set regex = New RegExp
    regex.Pattern = "{({{)*(\w+)(,-?\d+)?(:[^}]+)?}(}})*"
    regex.IgnoreCase = True
    regex.Global = True
    Set matches = regex.Execute(FormatString)
        
    'determine if valuesCount matches number of unique regex matches:
    Dim uniqueCount As Long
    Dim tempCSV As String
    For Each thisMatch In matches
        If Not StringType.Contains(tempCSV, thisMatch.SubMatches(1)) Then
            uniqueCount = uniqueCount + 1
            tempCSV = tempCSV & thisMatch.SubMatches(1) & ","
        End If
    Next
        
    'unique indices count must match valuesCount:
    If matches.Count > 0 And uniqueCount <> valuesCount Then _
       Err.Raise ERR_SOURCE, "Unique indices mismatch values count."
        
    If StringType.Contains(FormatString, "\\") Then _
                                               FormatString = Replace(FormatString, "\\", VBA.Strings.Chr$(27))
    
    If matches.Count = 0 And FormatString <> VBA.Constants.vbNullString And UBound(Values) = -1 Then
        'only formatString was specified: skip to checking escape sequences:
        returnValue = FormatString
        GoTo checkEscapes
    ElseIf UBound(Values) = -1 And matches.Count > 0 Then
        Err.Raise ERR_SOURCE, "Format specifier has no value."
    End If
        
    returnValue = FormatString
    
    'dissect formatString:
        
    '@Ignore UseMeaningfulName, MultipleDeclarations
    Dim i As Long, v As String, p As String      'i: iterator; v: value; p: placeholder
    '@Ignore MultipleDeclarations
    Dim alignmentGroup As String, alignmentSpecifier As String
    '@Ignore MultipleDeclarations
    Dim FormattedValue As String, alignmentPadding As Long
        
    'iterate regex matches (each match is a placeholder):
    For i = 0 To matches.Count - 1
            
        'get the placeholder specified index:
        Set thisMatch = matches.Item(i)
        p = thisMatch.SubMatches(1)
            
        'if specified index (0-based) > uniqueCount (1-based), something's wrong:
        If p > uniqueCount - 1 Then _
           Err.Raise ERR_SOURCE, "Format specifier index out of bounds."
        v = Values(p)
            
        'get the alignment specifier if it is specified:
        alignmentGroup = thisMatch.SubMatches(2)
        If alignmentGroup <> VBA.Constants.vbNullString Then _
           alignmentSpecifier = VBA.Strings.Right$(alignmentGroup, VBA.Strings.LenB(alignmentGroup) / 2 - 1)
            
            
        'get the format specifier if it is specified:
        thisString = thisMatch.Value
        If StringType.Contains(thisString, ":") Then
                
            '@Ignore MultipleDeclarations
            Dim formatGroup As String, precisionSpecifier As Long
            '@Ignore MultipleDeclarations
            Dim formatSpecifier As String, precisionString As String
                
            'get the string between ":" and "}":
            formatGroup = VBA.Strings.Mid$(thisString, VBA.Strings.InStr(1, thisString, ":") + 1, (VBA.Strings.LenB(thisString) / 2) - 2)
            formatGroup = VBA.Strings.Left$(formatGroup, VBA.Strings.LenB(formatGroup) / 2 - 1)
                
            precisionString = VBA.Strings.Right$(formatGroup, VBA.Strings.LenB(formatGroup) / 2 - 1)
            formatSpecifier = VBA.Strings.Mid$(thisString, VBA.Strings.InStr(1, thisString, ":") + 1, 1)
                                
            'applicable formatting depends on the type of the value (yes, GOTO!!):
            If VBA.Information.TypeName(Values(p)) = "Date" Then GoTo DateTimeFormatSpecifiers
            If v = VBA.Constants.vbNullString Then GoTo ApplyStringFormat
                
            'NumberFormatSpecifiers:
            If precisionString <> VBA.Constants.vbNullString And Not IsNumeric(precisionString) Then _
                                                                                                Err.Raise Exeption, _
                                                                                                ERR_SOURCE, ERR_MSG_INVALID_formatString
                
            If precisionString = VBA.Constants.vbNullString Then precisionString = 0
                
            Select Case formatSpecifier
                
            Case "C", "c"                        'CURRENCY format, formats string as currency.
                'Precision specifier determines number of decimal digits.
                'This implementation ignores regional settings
                '(hard-coded group separator, decimal separator and currency sign).
                    
                precisionSpecifier = VBA.Conversion.CInt(precisionString)
                thisFormat = "#,##0.00$"
                    
                If VBA.Strings.LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                    'if a non-zero precision is specified...
                    thisFormat = Replace$(thisFormat, ".00", "." & VBA.String$(precisionString, VBA.Strings.Chr$(48)))
                ElseIf VBA.Strings.LenB(formatGroup) > 2 And precisionSpecifier = 0 Then
                    thisFormat = Replace$(thisFormat, ".00", VBA.Constants.vbNullString)
                End If
                    
                    
            Case "D", "d"                        'DECIMAL format, formats string as Long number.
                'Precision specifier determines number of digits in returned string.
                    
                    
                precisionSpecifier = VBA.Conversion.CInt(precisionString)
                thisFormat = "0"
                thisFormat = VBA.Strings.Right$(VBA.String$(precisionSpecifier, "0") & thisFormat, _
                                                IIf(precisionSpecifier = 0, Len(thisFormat), precisionSpecifier))
                    
                    
            Case "E", "e"                        'EXPONENTIAL NOTATION format (aka "Scientific Notation")
                'Precision specifier determines number of decimals in returned string.
                'This implementation ignores regional settings'
                '(hard-coded decimal separator).
                    
                    
                precisionSpecifier = VBA.Conversion.CInt(precisionString)
                thisFormat = "0.00000#" & formatSpecifier & "-#" 'defaults to 6 decimals
                    
                If VBA.Strings.LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                    'if a non-zero precision is specified...
                    thisFormat = "0." & VBA.String$(precisionSpecifier - 1, VBA.Strings.Chr$(48)) & "#" & formatSpecifier & "-#"
                    
                ElseIf VBA.Strings.LenB(formatGroup) > 2 And precisionSpecifier = 0 Then
                    Err.Raise Exeption, _
                              ERR_SOURCE, ERR_MSG_INVALID_formatString
                End If
                    
                    
            Case "F", "f"                        'FIXED-POINT format
                'Precision specifier determines number of decimals in returned string.
                'This implementation ignores regional settings'
                '(hard-coded decimal separator).
                    
                precisionSpecifier = VBA.Conversion.CInt(precisionString)
                thisFormat = "0"
                If VBA.Strings.LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                    'if a non-zero precision is specified...
                    thisFormat = (thisFormat & ".") & VBA.String$(precisionSpecifier, VBA.Strings.Chr$(48))
                Else
                    'no precision specified - default to 2 decimals:
                    thisFormat = "0.00"
                End If
                    
                    
            Case "G", "g"                        'GENERAL format (recursive)
                'returns the shortest of either FIXED-POINT or SCIENTIFIC formats in case of a Double.
                'returns DECIMAL format in case of a Long or Long.
                    
                '@Ignore MultipleDeclarations
                Dim eNotation As String, ePower As Long, specifier As String
                precisionSpecifier = IIf(VBA.Conversion.CInt(precisionString) > 0, VBA.Conversion.CInt(precisionString), _
                                         IIf(StringType.Contains(v, "."), Len(v) - VBA.Strings.InStr(1, v, "."), 0))
                    
                'track character case of formatSpecifier:
                specifier = IIf(formatSpecifier = "G", "D", "d")
                    
                If VBA.Information.TypeName(Values(p)) = "Long" Or VBA.Information.TypeName(Values(p)) = "Long" Then
                    'Long types: use {0:D} (recursive call):
                    FormattedValue = StringType.Format("{0:" & specifier & "}", Values(p))
                    
                ElseIf VBA.Information.TypeName(Values(p)) = "Double" Then
                    'Non-Long types: use {0:E}
                    specifier = IIf(formatSpecifier = "G", "E", "e")
                        
                    'evaluate the exponential notation value (recursive call):
                    eNotation = StringType.Format("{0:" & specifier & "}", v)
                        
                    'get the power of eNotation:
                    ePower = VBA.Strings.Mid$(eNotation, VBA.Strings.InStr(1, VBA.UCase$(eNotation), "E-") + 1, _
                                              Len(eNotation) - VBA.Strings.InStr(1, VBA.UCase$(eNotation), "E-"))
                        
                    If ePower > -5 And Abs(ePower) < precisionSpecifier Then
                        'use {0:F} when ePower > -5 and abs(ePower) < precisionSpecifier:
                        'evaluate the floating-point value (recursive call):
                        specifier = IIf(formatSpecifier = "G", "F", "f")
                        FormattedValue = StringType.Format("{0:" & formatSpecifier & _
                                                           IIf(precisionSpecifier <> 0, precisionString, VBA.Constants.vbNullString) & "}", Values(p))
                    Else
                        'fallback to {0:E} if previous rule didn't apply:
                        FormattedValue = eNotation
                    End If
                        
                End If
                    
                GoTo AlignFormattedValue         'Skip the "ApplyStringFormat" step, it's applied already.
                    
                    
            Case "N", "n"                        'NUMERIC format, formats string as an Long or decimal number.
                'Precision specifier determines number of decimal digits.
                'This implementation ignores regional settings'
                '(hard-coded group and decimal separators).
                    
                precisionSpecifier = VBA.Conversion.CInt(precisionString)
                If VBA.Strings.LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                    'if a non-zero precision is specified...
                    thisFormat = "#,##0"
                    thisFormat = (thisFormat & ".") & VBA.String$(precisionSpecifier, VBA.Strings.Chr$(48))
                        
                Else                             'only the "D" is specified
                    thisFormat = "#,##0"
                End If
                    
                    
            Case "P", "p"                        'PERCENT format. Formats string as a percentage.
                'Value is multiplied by 100 and displayed with a percent symbol.
                'Precision specifier determines number of decimal digits.
                    
                thisFormat = "#,##0%"
                precisionSpecifier = VBA.Conversion.CInt(precisionString)
                If VBA.Strings.LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                    'if a non-zero precision is specified...
                    thisFormat = "#,##0"
                    thisFormat = (thisFormat & ".") & VBA.String$(precisionSpecifier, VBA.Strings.Chr$(48))
                        
                Else                             'only the "P" is specified
                    thisFormat = "#,##0"
                End If
                    
                'Append the percentage sign to the format string:
                thisFormat = thisFormat & "%"
                    
                    
            Case "R", "r"                        'ROUND-TRIP format (a string that can round-trip to an identical number)
                'example: ?StringFormat("{0:R}", 0.0000000001141596325677345362656)
                '         ...returns "0.000000000114159632567735"
                    
                'convert value to a Double (chop off overflow digits):
                v = CDbl(v)
                    
                    
            Case "X", "x"                        'HEX format. Formats a string as a Hexadecimal value.
                'Precision specifier determines number of total digits.
                'Returned string is prefixed with "&H" to specify Hex.
                    
                v = VBA.Hex$(v)
                precisionSpecifier = VBA.Conversion.CInt(precisionString)
                    
                If VBA.Strings.LenB(precisionString) > 0 Then 'precision here stands for left padding
                    v = VBA.Strings.Right$(VBA.String$(precisionSpecifier, "0") & v, IIf(precisionSpecifier = 0, Len(v), precisionSpecifier))
                End If
                    
                'add C# hex specifier, apply specified casing:
                '(VB6 hex specifier would cause Format() to reverse the formatting):
                v = "0x" & IIf(formatSpecifier = "X", VBA.UCase$(v), VBA.LCase$(v))
                    
                    
            Case Else
                Err.Raise Exeption, _
                          ERR_SOURCE, ERR_MSG_INVALID_formatString
            End Select
                
            GoTo ApplyStringFormat
                
                
DateTimeFormatSpecifiers:
            Select Case formatSpecifier
                    
            Case "c", "C"                        'CUSTOM date/time format
                'let VB Format() parse precision specifier as is:
                thisFormat = precisionString
                    
            Case "d"                             'SHORT DATE format
                thisFormat = "ddddd"
                        
            Case "D"                             'LONG DATE format
                thisFormat = "dddddd"
                        
            Case "f"                             'FULL DATE format (short)
                thisFormat = "dddddd h:mm AM/PM"
                    
            Case "F"                             'FULL DATE format (long)
                thisFormat = "dddddd ttttt"
                    
            Case "g"
                thisFormat = "ddddd hh:mm AM/PM"
                        
            Case "G"
                thisFormat = "ddddd ttttt"
                        
            Case "s"                             'SORTABLE DATETIME format
                thisFormat = "yyyy-mm-ddThh:mm:ss"
                    
            Case "t"                             'SHORT TIME format
                thisFormat = "hh:mm AM/PM"
                    
            Case "T"                             'LONG TIME format
                thisFormat = "ttttt"
                    
            Case Else
                Err.Raise Exeption, _
                          ERR_SOURCE, ERR_MSG_INVALID_formatString
            End Select
            GoTo ApplyStringFormat
                
        End If
            
            
ApplyStringFormat:
        'apply computed format string:
        FormattedValue = VBA.Strings.Format$(v, thisFormat)
            
            
AlignFormattedValue:
        'apply specified alignment specifier:
        If alignmentSpecifier <> VBA.Constants.vbNullString Then
            
            alignmentPadding = Abs(VBA.Conversion.CInt(alignmentSpecifier))
            If VBA.Conversion.CInt(alignmentSpecifier) < 0 Then
                'negative: left-justified alignment
                If alignmentPadding - Len(FormattedValue) > 0 Then _
                                                          FormattedValue = FormattedValue & _
                                                          VBA.String$(alignmentPadding - Len(FormattedValue), PADDING_CHAR)
            Else
                'positive: right-justified alignment
                If alignmentPadding - Len(FormattedValue) > 0 Then _
                                                          FormattedValue = VBA.String$(alignmentPadding - Len(FormattedValue), PADDING_CHAR) & FormattedValue
            End If
        End If
            
        'Replace C# hex specifier with VB6 hex specifier:
        If StringType.Contains(FormattedValue, "0x") Then FormattedValue = Replace$(FormattedValue, "0x", "&H")
            
        'replace all occurrences of placeholder {i} with their formatted values:
        returnValue = Replace(returnValue, thisString, FormattedValue, Count:=1)
            
        'reset before reiterating:
        thisFormat = VBA.Constants.vbNullString
    Next
        
        
checkEscapes:
    'if there's no more backslashes, don't bother checking for the rest:
    If Not StringType.Contains(returnValue, "\") Then GoTo normalExit
        
    Dim Escape As EscapeSequence
    For i = 0 To Escapes.Count - 1
        Set Escape = Escapes.Item(CStr(i))
        If StringType.Contains(returnValue, Escape.EscapeString) Then _
                                                                 returnValue = Replace(returnValue, Escape.EscapeString, Escape.ReplacementString)
        
        If Not StringType.Contains(returnValue, "\") Then _
                                                     GoTo normalExit
    Next
        
    'replace "ASCII (oct)" escape sequence
    Set regex = New RegExp
    regex.Pattern = "\\(\d{3})"
    regex.IgnoreCase = True
    regex.Global = True
    Set matches = regex.Execute(FormatString)
        
    Dim Char As Long
    If matches.Count <> 0 Then
        For Each thisMatch In matches
            p = thisMatch.SubMatches(0)
            '"p" contains the octal number representing the ASCII code we're after:
            p = "&O" & p                         'prepend octal prefix
            Char = CLng(p)
            returnValue = Replace(returnValue, thisMatch.Value, VBA.Strings.Chr$(Char))
        Next
    End If
        
    'if there's no more backslashes, don't bother checking for the rest:
    If Not StringType.Contains("\", returnValue) Then GoTo normalExit
        
    'replace "ASCII (hex)" escape sequence
    Set regex = New RegExp
    regex.Pattern = "\\x(\w{2})"
    regex.IgnoreCase = True
    regex.Global = True
    Set matches = regex.Execute(FormatString)
        
    If matches.Count <> 0 Then
        For Each thisMatch In matches
            p = thisMatch.SubMatches(0)
            '"p" contains the hex value representing the ASCII code we're after:
            p = "&H" & p                         'prepend hex prefix
            Char = CLng(p)
            returnValue = Replace(returnValue, thisMatch.Value, VBA.Strings.Chr$(Char))
        Next
    End If

normalExit:
    If StringType.Contains(returnValue, VBA.Strings.Chr$(27)) Then returnValue = Replace(returnValue, VBA.Strings.Chr$(27), "\")
    Format = returnValue
   
End Function

'@Description "Renvoie TRUE si stringSource contient findText (case insensitive by default)."
Public Function Contains(ByVal stringSource As String, ByVal FindText As String, _
                         Optional ByVal caseSensitive As Boolean = False) As Boolean
Attribute Contains.VB_Description = "Renvoie TRUE si stringSource contient findText (case insensitive by default)."
        
    Dim compareMethod As VbCompareMethod
    
    If caseSensitive Then
        compareMethod = vbBinaryCompare
    Else
        compareMethod = vbTextCompare
    End If
    Contains = (VBA.Strings.InStr(1, stringSource, FindText, compareMethod) <> 0)
    
End Function

'@Description "Renvoie TRUE si stringSource contient n'importe laquelle des valeurs specifiees."
Public Function ContainsAny(ByVal stringSource As String, ByVal caseSensitive As Boolean, _
                            ParamArray findStrings() As Variant) As Boolean
Attribute ContainsAny.VB_Description = "Renvoie TRUE si stringSource contient n'importe laquelle des valeurs specifiees."
    
    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = LBound(findStrings) To UBound(findStrings)
        Dim Find As String: Find = CStr(findStrings(i))
        Dim Found As Boolean: Found = Contains(stringSource, Find, caseSensitive)
        If Found Then Exit For
    Next
    ContainsAny = Found
    
End Function

'@Description "Renvoie TRUE si stringSource égale n'importe laquelle des valeurs specifiées."
Public Function MatchesAny(ByVal stringSource As String, ParamArray findStrings() As Variant) As Boolean
Attribute MatchesAny.VB_Description = "Renvoie TRUE si stringSource égale n'importe laquelle des valeurs specifiées."
    
    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = LBound(findStrings) To UBound(findStrings)
    
        Dim Find As String
        Find = CStr(findStrings(i))
        Dim Found As Boolean
        Found = (stringSource = Find)
        
        If Found Then Exit For
    Next
    MatchesAny = Found
    
End Function

'@Description "Renvoie TRUE si stringSource égale toutes les valeurs specifiées."
Public Function MatchesAll(ByVal stringSource As String, ParamArray findStrings() As Variant) As Boolean
Attribute MatchesAll.VB_Description = "Renvoie TRUE si stringSource égale toutes les valeurs specifiées."
    
    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = LBound(findStrings) To UBound(findStrings)
        Dim Find As String
        Find = CStr(findStrings(i))
        Dim Match As Boolean
        Match = (stringSource = Find)
        If Not Match Then Exit For
    Next
    MatchesAll = Match
    
End Function

'@Description "Renvoie TRUE si stringSource débute par findText (case sensitive)."
Public Function StartsWith(ByVal FindText As String, ByVal stringSource As String, _
                           Optional ByVal caseSensitive As Boolean = True) As Boolean
Attribute StartsWith.VB_Description = "Renvoie TRUE si stringSource débute par findText (case sensitive)."
    
    If Not caseSensitive Then
        Dim localFindText As String: localFindText = FindText
        localFindText = VBA.LCase$(localFindText)
        Dim localStringSource As String: localStringSource = stringSource
        localStringSource = VBA.LCase$(localStringSource)
    End If
    StartsWith = (VBA.Strings.Left$(localStringSource, VBA.Strings.LenB(localFindText) / 2) = localFindText)

End Function

Public Function StartsWithAny(ByVal stringSource As String, ByVal caseSensitive As Boolean, _
                              ParamArray findStrings() As Variant) As Boolean
    
    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = LBound(findStrings) To UBound(findStrings)
        Dim Find As String: Find = CStr(findStrings(i))
        Dim Found As Boolean: Found = StartsWith(Find, stringSource, caseSensitive)
        If Found Then Exit For
    Next
    StartsWithAny = Found
    
End Function

'@Description "Renvoie TRUE si stringSource se termine par findText (case sensitive)."
Public Function EndsWith(ByVal FindText As String, ByVal stringSource As String, _
                         Optional ByVal caseSensitive As Boolean = True) As Boolean
Attribute EndsWith.VB_Description = "Renvoie TRUE si stringSource se termine par findText (case sensitive)."
        
    If Not caseSensitive Then
        Dim localStringSource As String: localStringSource = stringSource
        localStringSource = VBA.LCase$(localStringSource)
        Dim localFindText As String: localFindText = FindText
        localFindText = VBA.LCase$(localFindText)
    End If
    EndsWith = (VBA.Strings.Right$(localStringSource, VBA.Strings.LenB(localFindText) / 2) = localFindText)

End Function

Public Function EndsWithAny(ByVal stringSource As String, ByVal caseSensitive As Boolean, _
                            ParamArray findStrings() As Variant) As Boolean
    
    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = LBound(findStrings) To UBound(findStrings)
        Dim Find As String: Find = CStr(findStrings(i))
        Dim Found As Boolean: Found = EndsWith(Find, stringSource, caseSensitive)
        If Found Then Exit For
    Next
    EndsWithAny = Found
    
End Function

Public Function Coalesce(ParamArray parms() As Variant) As Variant
    '@Ignore UseMeaningfulName
    Dim i As Long
    For i = 0 To UBound(parms)
        Dim currentParm As Variant
        currentParm = parms(i)
        Coalesce = currentParm
        If Not IsNull(currentParm) Then
            Exit Function
        End If
    Next
End Function
