VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Services"
Option Explicit

Private objRegEx As Object

'Following function is to filter table data by checking with different formats of data!
'inputValue would be table's Cell Data on which we have test if that cell meets criteria or not
'checkType would give idea that what kind of checking this is! either we are cheking two numbers, two string or dates?
'checkValue1 and checkValue2 are criteria values! if only one of them is mentioned then it is simply binary checking
'if both are mentioned then presently checking if the date is between two mentioned dates or not?
'the same way this validation can be evolved with multiple check types like, numerical ranges etc.etc.

Public Function IsValidToInclude(ByVal inputValue As Variant, ByVal checkType As ValidationCheckTypes, Optional ByVal checkValue1 As Variant, Optional ByVal checkValue2 As Variant) As Boolean
    IsValidToInclude = False
    'Check 1
    If VBA.IsNumeric(inputValue) Then
        If inputValue = 0 Then IsValidToInclude = False Else IsValidToInclude = True
    Else
        If inputValue = vbNullString Then IsValidToInclude = False Else IsValidToInclude = True
    End If
    If IsValidToInclude = False Then Exit Function
    'Check 2
    If VBA.IsNumeric(checkValue1) Then
        If checkValue1 = 0 Then IsValidToInclude = True Else IsValidToInclude = False
    Else
        If checkValue1 = vbNullString Then IsValidToInclude = True Else IsValidToInclude = False
    End If
    If IsValidToInclude = True Then Exit Function
    'Check 3
    Select Case checkType
        Case ValidationCheckTypes.TYPE_DATEBETWEENRANGE
            If VBA.CDate(inputValue) >= VBA.CDate(checkValue1) And _
                VBA.CDate(inputValue) <= VBA.CDate(checkValue2) Then
                IsValidToInclude = True
            Else
                IsValidToInclude = False
            End If
        Case ValidationCheckTypes.TYPE_STRINGMATCH
            If VBA.IsNumeric(checkValue1) Then
                IsValidToInclude = (Val(inputValue) = Val(checkValue1))
            Else
                IsValidToInclude = (inputValue = checkValue1)
            End If
    End Select
End Function

'Check if the date is between the start and end range dates or not?
Public Function IsValidDateInRange(ByVal inputValue As String, ByVal startDate As Date, ByVal endDate As Date) As Boolean
    Dim TargetDate As Date
    TargetDate = VBA.CDate(GetDateFromGermanDateString(inputValue))
    If TargetDate >= startDate And TargetDate <= endDate Then
        IsValidDateInRange = True
    Else
        IsValidDateInRange = False
    End If
End Function

'Check if the value is Valid German accepted date format or not?
'[10.02.2009] [10022009] [31.12.2022]
Public Function IsValidGermanFormatDate(ByVal inputValue As String) As Boolean
    'use part strings to get if the string is date or not?
    If GetDateFromGermanDateString(inputValue) <> vbNullString Then
        IsValidGermanFormatDate = True
    Else
        IsValidGermanFormatDate = False
    End If
End Function

'Check if Value is between Numerical Range or not?
Public Function IsBetweenNumericalRange(ByVal inputValue As String, _
                                        ByVal minValue As Long, _
                                        ByVal maxvalue As Long) As Boolean
    Dim result As Boolean
    
    On Error GoTo CleanFail
    
    'owerflow error is rased if len(inputValue) > len(maxvalue)
    result = False 'I have habit of not depending on default value of any type of variables! I explicitly assign value to it!
    If VBA.CLng(inputValue) > minValue And VBA.CLng(inputValue) <= maxvalue Then
        result = True
    Else
        result = False
    End If
    
CleanExit:
    IsBetweenNumericalRange = result
    Exit Function

CleanFail:
    MsgBox "ErrNumber:" & Err.Number & " Err.Description: " & Err.Description, vbCritical
    Resume CleanExit
    Resume
    
End Function

'if All goes well for the form
Public Function IsEverythingOkay(ByVal ArrayOfCheckFieldsStatus As Variant, _
                                ByVal ArrayOfTableFields As Variant, _
                                ByVal FirstColIndex As Long, _
                                ByVal LastColIndex As Long) As Variant
    Dim flag As Boolean
    Dim tmp As String
    Dim I As Long
    flag = False
    tmp = vbNullString
    'Checking each properties validation state
    For I = FirstColIndex To LastColIndex Step 1
        If ArrayOfCheckFieldsStatus(I) = False Then
            flag = False
            Exit For
        Else
            flag = True
        End If
    Next I
    'Moment of Truth
    If flag = True Then
        IsEverythingOkay = True
    Else
        'Checking each properties validation state
        For I = FirstColIndex To LastColIndex Step 1
            If ArrayOfCheckFieldsStatus(I) = False Then
                tmp = tmp & BULLET_LISTITEM & ArrayOfTableFields(I - 1) & vbNewLine
            End If
        Next I
        IsEverythingOkay = "Kindly fill the following fields(s) properly to proceed further!" & vbNewLine & vbNewLine & tmp
    End If
End Function

'check if the value is empty or not?
Public Function IsNotNullValue(ByVal inputValue As Variant, ByVal inputType As String) As Boolean
    Select Case inputType
        Case "Byte", "Integer", "Long", "Single", "Double"
            If inputValue = 0 Then
                IsNotNullValue = False
            Else
                IsNotNullValue = True
            End If
        Case "String"
            If inputValue = vbNullString Then
                IsNotNullValue = False
            Else
                IsNotNullValue = True
            End If
        Case "Variant"
            If VBA.IsNull(inputValue) Then
                IsNotNullValue = False
            Else
                IsNotNullValue = True
            End If
    End Select
End Function

'Check if string has crossed max limit or not?
Public Function IsValidStringLengthUnderLimit(ByVal inputString As Variant, ByVal maxLength As Byte) As Boolean
    Dim checklength As Byte
    checklength = maxLength
    'Following logic is to pass currency
    If VBA.IsNumeric(inputString) Then
        If VBA.InStr(1, inputString, ",") > 0 Then checklength = checklength + 1
        If VBA.InStr(1, inputString, ".") > 0 Then checklength = checklength + 1
    End If
    'Main Logic
    If VBA.Len(inputString) <= checklength Then
        IsValidStringLengthUnderLimit = True
    Else
        IsValidStringLengthUnderLimit = False
    End If
End Function

'Check for the Valid Length
Public Function IsValidStringLength(ByVal inputString As Variant, ByVal checklength As Byte) As Boolean
    If VBA.Len(inputString) >= checklength Then
        IsValidStringLength = True
    Else
        IsValidStringLength = False
    End If
End Function

'check if the user is authorized or not?
Public Function IsUserAuthorized(ByVal dataTable As Variant, ByVal userName As String, ByVal password As String, ByVal userCol As Long, ByVal passCol As Long) As Variant
    Dim I As Long
    Dim userFlag As Boolean
    Dim passFlag As Boolean
    userFlag = False
    passFlag = False
    'validation for initial check
    If userName = vbNullString And password = vbNullString Then
        IsUserAuthorized = "Username and Passwrod fields are empty! Kindly fill in to Login"
        Exit Function
    ElseIf userName = vbNullString And password <> vbNullString Then
        IsUserAuthorized = "Username field is empty! Kindly fill in to Login."
        Exit Function
    ElseIf userName <> vbNullString And password = vbNullString Then
        IsUserAuthorized = "Password field is empty! Kindly fill in to Login."
        Exit Function
    End If
    'if validation check confirm then go for the next process
    For I = 1 To UBound(dataTable, 1) Step 1
        If dataTable(I, userCol) = userName Then
        userFlag = True
            If dataTable(I, passCol) = password Then
                passFlag = True
            End If
        End If
    Next I
    'conclusion
    If userFlag = False Then
        IsUserAuthorized = "Invalid Username!"
    Else
        If passFlag = False Then
            IsUserAuthorized = "Invalid Password!"
        Else
            IsUserAuthorized = True
        End If
    End If
End Function

'following function will compare the strings

Public Function ComapreTwoStrings(ByVal String1 As String, ByVal String2 As String) As Boolean
    If String1 = String2 Then
        ComapreTwoStrings = True
    Else
        ComapreTwoStrings = False
    End If
End Function

'Following function will verify if the password string contains following characters or not?

Public Function IsValidPasswordString(ByVal passwordString As String) As Boolean
    'Variable declarations
    Dim txt                     As String
    Dim passwordCharacters()    As String
    Dim passwordCharacter       As Variant
    Dim testCapital             As Boolean
    Dim testLower               As Boolean
    Dim testNumber              As Boolean
    Dim testSpecialCharacter    As Boolean
    'Check for the Minimum Length Criteria
    If VBA.Len(passwordString) >= 8 Then
        IsValidPasswordString = True
    Else
        IsValidPasswordString = False
        Exit Function
    End If
    'Get characters in array
    txt = VBA.StrConv(passwordString, vbUnicode)
    passwordCharacters = VBA.Split(VBA.Left(txt, Len(txt) - 1), VBA.Chr$(0))
    'Get testified
    'Criterias of checking
        'Minimum 8 characters of Length
        'At least One Capital
        'At least One Small Character
        'At least One Numeric
        'At least One Special Character
    For Each passwordCharacter In passwordCharacters
        Select Case VBA.AscW(passwordCharacter)
            'A-Z : 65-90
            Case 65 To 90
                testCapital = True
            'a-z
            Case 97 To 122
                testLower = True
            '0-9
            Case 48 To 57
                testNumber = True
            'Special Characters
            Case 33 To 47, 58 To 64, 91 To 96, 123 To 126
                testSpecialCharacter = True
        End Select
    Next passwordCharacter
    'final conclusion
    If (testCapital And testLower And testNumber And testSpecialCharacter) Then
        IsValidPasswordString = True
    Else
        IsValidPasswordString = False
    End If
End Function

'Helping Functions

'This will parse the German formatted date string into workable format
Public Function GetDateFromGermanDateString(ByVal inputValue As String) As String
    Dim dayPart As Long
    Dim MonthPart As Long
    Dim YearPart As Long
    Dim tmp As String
    'clean input string so we only left with 8 characters long length numerical value
    tmp = CleanString(inputValue)
    inputValue = tmp
    'check for [10022009]DDMMYYYY
    If VBA.Len(inputValue) = 8 And VBA.IsNumeric(inputValue) Then
        'you are here because inputValue has right 8 characters that suits to DDMMYYYY
        dayPart = VBA.CLng(VBA.Left(inputValue, 2))
        YearPart = VBA.CLng(VBA.Right(inputValue, 4))
        MonthPart = VBA.CLng(VBA.Mid(inputValue, 3, 2))
        GetDateFromGermanDateString = VBA.Format(VBA.DateSerial(YearPart, MonthPart, dayPart), DATEFORMAT_BACKEND)
    Else
        GetDateFromGermanDateString = vbNullString
    End If
End Function

'this function uses regexp library to remove unwanted characters from the string and it only keeps numericals!
Private Function CleanString(ByVal inputValue As String) As String
    With objRegEx
        .Global = True
        .Pattern = "[^\d]+"
        CleanString = .Replace(inputValue, vbNullString)
    End With
End Function

'This function will reset each items of array to false or reset it

Public Sub ResetSingleDimensionalArray(ByRef arr As Variant, Optional resetValue As Boolean = False)
    Dim I As Long
    Dim FirstItem As Long
    Dim LastItem As Long
    FirstItem = LBound(arr)
    LastItem = UBound(arr)
    For I = FirstItem To LastItem Step 1
        arr(I) = resetValue
    Next I
End Sub

'Class Events

Private Sub Class_Initialize()
    Set objRegEx = CreateObject("vbscript.regexp")
End Sub

Private Sub Class_Terminate()
    Set objRegEx = Nothing
End Sub
