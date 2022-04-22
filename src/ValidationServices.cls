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

'Check for the Valid Length
Public Function IsValidStringLength(ByVal inputstring As Variant, ByVal checkLength As Byte) As Variant
    If VBA.Len(inputstring) >= checkLength Then
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

