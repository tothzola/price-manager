VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@PredeclaredId
'@Folder("MVP.Model")
Option Explicit

Private Const SEPERATOR  As String = "<ITEM>"

Private Type TAppMethods
    ObjRegEx As Object
    Algo As IHashAlgorithm
End Type

Private This As TAppMethods

'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As AppMethods
Attribute Class.VB_Description = "Returns class reference"
    Set Class = AppMethods
End Property

Private Sub Class_Initialize()
    Guard.Singleton Me
End Sub

Public Property Get ObjRegEx() As Object
    If This.ObjRegEx Is Nothing Then
        Set This.ObjRegEx = CreateObject("vbscript.regexp")
    End If
    Set ObjRegEx = This.ObjRegEx
    
End Property

Public Property Get Algo() As IHashAlgorithm
    If This.Algo Is Nothing Then
        Set This.Algo = SHA256Managed.Create()
    End If
    Set Algo = This.Algo
    
End Property

'@Description("Function return (6) = YES or (7) = NO depending on user selection")
Public Function UserNotifierConfirmation(ByVal msgPrompt As String, ByVal msgTitel As String) As Long

    UserNotifierConfirmation = VBA.Interaction.MsgBox( _
                                        Prompt:=msgPrompt, _
                                        Buttons:=VBA.vbQuestion + VBA.vbYesNo + VBA.vbDefaultButton1, _
                                        Title:=msgTitel)

End Function

'@Description("Create a Randomized String of Characters, SOURCE: www.TheSpreadsheetGuru.com/the-code-vault")
Public Function RandomString(ByVal Length As Integer) As String
Attribute RandomString.VB_Description = "Create a Randomized String of Characters, SOURCE: www.TheSpreadsheetGuru.com/the-code-vault"

    Dim CharacterBank   As Variant
    Dim x               As Long
    Dim str             As String
    
    'Test Length Input
    If Length < 1 Then
        MsgBox "Length variable must be greater than 0"
        Exit Function
    End If
    
    CharacterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
    "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
    "y", "z", "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "!", "@", _
    "#", "$", "%", "^", "&", "*", "A", "B", "C", "D", "E", "F", "G", "H", _
    "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
    "W", "X", "Y", "Z")
    
    
    'Randomly Select Characters One-by-One
    For x = 1 To Length
        VBA.Randomize
        str = str & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * VBA.Rnd + LBound(CharacterBank)))
    Next x
    
    'Output Randomly Generated String
    RandomString = str

End Function

'@Description("This function will crypt the string value")
Public Function CryptString(ByVal InputValue As String) As String
Attribute CryptString.VB_Description = "This function will crypt the string value"
    Guard.EmptyString InputValue
    
    'cript password
    Dim cryptedString As String
    If Not Algo Is Nothing Then cryptedString = Algo.ComputeHash(InputValue)
    
    CryptString = cryptedString
    
End Function

'@Description("This function will reset each items of array to false or reset it")
Public Sub ResetSingleDimensionalArray(ByRef Arr As Variant, Optional resetValue As Boolean = False)
Attribute ResetSingleDimensionalArray.VB_Description = "This function will reset each items of array to false or reset it"
    Dim i As Long
    Dim FirstItem As Long
    Dim LastItem As Long
    FirstItem = LBound(Arr)
    LastItem = UBound(Arr)
    For i = FirstItem To LastItem
        Arr(i) = resetValue
    Next i
End Sub

'@Description("Following function will tell if the string is currency or not? _
                Basically, in this app, if the number is representing the currency _
                then it must have decimal seperator! So we have to check for the _
                decimal seperator's availability only. If it is Available at the _
                third from the right side of the number then it must be a currency")
Public Function IsCurrency(ByVal InputValue As String) As Boolean
    
    Dim commaChar As String
    commaChar = VBA.Left(VBA.Right(InputValue, 3), 1)
    IsCurrency = (commaChar = Application.International(xlDecimalSeparator))
    
End Function

'@Description("Following function will return new array by filtering it from one it's column")
Public Function FilterListWithSingleCondition(ByVal ContextTable As Variant, ByVal ColumnIndex As Long, ByVal condition As Variant) As Variant
Attribute FilterListWithSingleCondition.VB_Description = "Following function will return new array by filtering it from one it's column"
    'Driving Paramters
    Const CONSIDERHEADERINDEX   As String = "1" & SEPERATOR
    Dim Records()               As String
    Dim Record                  As Variant
    Dim tmp                     As Variant
    Dim MatchedRecords          As String
    Dim TotalMatchingRows       As Long
    Dim firstRow                As Long
    Dim LastRow                 As Long
    Dim FirstCol                As Long
    Dim LastCol                 As Long
    Dim i                       As Long
    Dim j                       As Long
    'Get Driving parameters
    firstRow = LBound(ContextTable, 1)
    LastRow = UBound(ContextTable, 1)
    FirstCol = LBound(ContextTable, 2)
    LastCol = UBound(ContextTable, 2)
    ReDim tmp(1 To 1, 1 To LastCol)
    'Get matching Rows!
    MatchedRecords = CONSIDERHEADERINDEX
    TotalMatchingRows = 0
    For i = firstRow To LastRow
        If ContextTable(i, ColumnIndex) = condition Then
            MatchedRecords = MatchedRecords & i & SEPERATOR
            TotalMatchingRows = TotalMatchingRows + 1
        End If
    Next i
    'remove last seperator from the matchedrecords variable
    If MatchedRecords = CONSIDERHEADERINDEX Then
        For j = FirstCol To LastCol
            tmp(1, j) = ContextTable(1, j)
        Next j
        FilterListWithSingleCondition = tmp
        Exit Function
    Else
        MatchedRecords = VBA.Left$(MatchedRecords, VBA.Len(MatchedRecords) - VBA.Len(SEPERATOR))
    End If
    'Redimentioning the tmp array
    ReDim tmp(firstRow To TotalMatchingRows + 1, FirstCol To LastCol)
    'Transferring Data
    Records = VBA.Split(MatchedRecords, SEPERATOR)
    j = 1
    For Each Record In Records
        For i = FirstCol To LastCol
            tmp(j, i) = ContextTable(Record, i)
        Next i
        j = j + 1
    Next Record
    'return value
    FilterListWithSingleCondition = tmp
End Function


'@Description("Following table will join Table 1 on Table 2. First Table 1 will be _
                populated on new array then new table will join thereafter. This is _
                Two Dimensional Arrays")
Public Function JoinTables(ByVal Table1 As Variant, ByVal Table2 As Variant) As Variant
    Dim TotalRowsTable1     As Long
    Dim TotalRowsTable2     As Long
    Dim TotalColumnsTable1  As Long
    Dim TotalColumnsTable2  As Long
    Dim FinalTableRows      As Long
    Dim FinalTableColumns   As Long
    Dim i                   As Long
    Dim j                   As Long
    Dim FinalTable          As Variant
    'Check if the table2 is Double dimensional or not?
    'if not then create it
    If NumberOfArrayDimensions(Table2) = 1 Then
        Dim tmp As Variant
        ReDim tmp(1 To 1, 1 To UBound(Table1, 2))
        For i = 1 To UBound(Table1, 2)
            tmp(1, i) = Table2(i)
        Next i
        Table2 = tmp
    End If
    'Everything is clear!
    TotalRowsTable1 = UBound(Table1, 1)
    TotalRowsTable2 = UBound(Table2, 1)
    TotalColumnsTable1 = UBound(Table1, 2)
    TotalColumnsTable2 = UBound(Table2, 2)
    If TotalColumnsTable1 <> TotalColumnsTable2 Then
        JoinTables = Table1
        Exit Function
    End If
    FinalTableRows = TotalRowsTable1 + TotalRowsTable2
    FinalTableColumns = TotalColumnsTable1
    ReDim FinalTable(1 To FinalTableRows, 1 To FinalTableColumns)
    For i = 1 To FinalTableRows
        For j = 1 To FinalTableColumns
            If i <= TotalRowsTable1 Then
                FinalTable(i, j) = Table1(i, j)
            Else
                FinalTable(i, j) = Table2(i - TotalRowsTable1, j)
            End If
        Next j
    Next i
    JoinTables = FinalTable
End Function

'@Description("Change text format to #,###.00 Currency Format")
Public Function FormatStringToCurrency(ByVal InputValue As String, Optional prefix As String = vbNullString) As String
Attribute FormatStringToCurrency.VB_Description = "Change text format to #,###.00 Currency Format"
    
    'Variable Declaration
    Dim IsDecimalSeperatorUsed  As Boolean
    Dim ObjRegEx                As Object
    Dim ThousandSeperator       As String
    Dim DecimalSeperator        As String
    Dim DecimalPart             As String
    Dim WholeNumberPart         As String
    Dim ArrDecimalNumber()      As String
    Dim tmp                     As String
    Dim LenInputValue           As Long
    Dim LenDecimalPart          As Long
    Dim i                       As Long
    Dim j                       As Long
        
    Const Error_NonNumericalInput As Long = 1
    
    'error handeling
    On Error GoTo EhhErr
    
    'Validation Checks
    
        'Check 1 : On null value
        If InputValue = vbNullString Then GoTo CleanExit
        
        'Check 2 : If Non-Numerical Value entered!
        If prefix <> vbNullString Then
            InputValue = VBA.Replace(InputValue, prefix & " ", vbNullString)
        End If
        If Not VBA.IsNumeric(InputValue) Then
            'Err.Raise _
                Number:=Error_NonNumericalInput + 513, _
                Description:="Non-Numerical character(s) are not allowed!"
            MsgBox "Non-Numerical character(s) are not allowed!", vbCritical, Title:=SIGN
            LogManager.Log ErrorLevel, "Error: " & VBA.Err.Number & " Non-Numerical character(s) are not allowed!"
        End If
        
    'Get Sepearator String
    ThousandSeperator = Application.International(xlThousandsSeparator)
    DecimalSeperator = Application.International(xlDecimalSeparator)
    InputValue = Me.TransformToAPPCompatibleCurrency(InputValue)
    
    'check if decimal seperator has been used or not ?
    If VBA.InStr(1, InputValue, DecimalSeperator) > 0 Then
        'Get number parts with Position shifter
        ArrDecimalNumber = VBA.Split(InputValue, DecimalSeperator)
        LenDecimalPart = VBA.Len(ArrDecimalNumber(1))
        If LenDecimalPart = 2 Then
            WholeNumberPart = ArrDecimalNumber(0)
            DecimalPart = ArrDecimalNumber(1)
        ElseIf LenDecimalPart = 3 Then
            If ArrDecimalNumber(0) = 0 Then
                WholeNumberPart = VBA.Left$(ArrDecimalNumber(1), 1)
            Else
                WholeNumberPart = ArrDecimalNumber(0) & VBA.Left$(ArrDecimalNumber(1), 1)
            End If
            DecimalPart = VBA.Right$(ArrDecimalNumber(1), 2)
        ElseIf LenDecimalPart = 1 Then
            FormatStringToCurrency = vbNullString
            GoTo CleanExit
        End If
    Else
        'Length of Input String
        LenInputValue = VBA.Len(InputValue)
        'Get whole number parts
        If LenInputValue <= 2 Then
            WholeNumberPart = 0
        Else
            WholeNumberPart = VBA.Left$(InputValue, LenInputValue - 2)
        End If
        'get decimal number part
        If LenInputValue > 1 Then
            DecimalPart = VBA.Right$(InputValue, 2)
        Else
            DecimalPart = "0" & InputValue
        End If
    End If
    
    'clean number parts
    Set ObjRegEx = CreateObject("vbscript.regexp")
    With ObjRegEx
        .Global = True
        .Pattern = "[^\d]+"
        WholeNumberPart = .Replace(WholeNumberPart, vbNullString)
        DecimalPart = .Replace(DecimalPart, vbNullString)
    End With
    
    'insert Thousand Seperator to whole number part!
    tmp = vbNullString
    If VBA.Len(WholeNumberPart) > 3 Then
        j = 1
        For i = VBA.Len(WholeNumberPart) To 1 Step -1
            If j = 3 Then
                tmp = ThousandSeperator & VBA.Mid$(WholeNumberPart, i, 1) & tmp
                j = 0
            Else
                tmp = VBA.Mid$(WholeNumberPart, i, 1) & tmp
            End If
            j = j + 1
        Next i
        'Clean whole number part
        If VBA.Left$(tmp, 1) = ThousandSeperator Then
            tmp = VBA.Right$(tmp, VBA.Len(tmp) - 1)
            tmp = VBA.Trim$(tmp)
        End If
        'Submission!
        WholeNumberPart = tmp
    End If
    
    'Joining String
    If prefix = vbNullString Then
        FormatStringToCurrency = WholeNumberPart & DecimalSeperator & DecimalPart
    Else
        FormatStringToCurrency = prefix & " " & WholeNumberPart & DecimalSeperator & DecimalPart
    End If
    
CleanExit:
    Set ObjRegEx = Nothing
    Exit Function
    
EhhErr:
    MsgBox "Error: " & Err.Number & vbNewLine & vbNewLine & Err.Description, vbCritical, SIGN
    FormatStringToCurrency = vbNullString
    Err.Clear
    LogManager.Log ErrorLevel, "Error: " & Err.Number & vbTab & Err.Description
    Resume CleanExit
    
End Function


'@Description("Following function will convert currency to Application Compatible currency format! _
                this is very interesting because, We should always keep database column format in _
                single culture format. Like, en_US format, and we always get data from the database _
                in decimal seperator dot format. So if in case our user is german, than that currency will _
                be invalid string and hence, we have to make few changes by replacing dot with comma to make _
                it compitable with App format. Rest of the work will be carried by then FormatStringToCurrency Function.")
Public Function TransformToAPPCompatibleCurrency(ByVal InputValue As String) As String
    If Application.International(xlDecimalSeparator) = "," Then
        InputValue = VBA.Replace(InputValue, ".", ",")
    End If
    TransformToAPPCompatibleCurrency = InputValue
End Function

'@Description("Following function will convert currency to DB Compatible currency format! _
                this is very interesting because, We should always keep database column format in _
                single culture format. Like, en_US format, so decimal input will be always seperated with _
                dot! so, and why we need this? because german inputs considers comma as decimal seperator, _
                so we have to transform the currency string into the DB compitable string and then we can only _
                make it workable.")
Public Function TransformToDBCompatibleCurrency(ByVal InputValue As String) As String
    If Application.International(xlDecimalSeparator) = "." Then
        InputValue = VBA.Replace(InputValue, ",", vbNullString)
    Else
        InputValue = VBA.Replace(VBA.Replace(VBA.Replace(InputValue, ",", "#"), ".", vbNullString), "#", ".")
    End If
    TransformToDBCompatibleCurrency = InputValue
End Function

'@Description("Following function will pick unique list from the list!")
Public Function GetUniqueListFromTableColumn(ByVal ContextTable As Variant, ByVal TargetColumn As Long) As Variant
Attribute GetUniqueListFromTableColumn.VB_Description = "Following function will pick unique list from the list!"
    
    'Driving Paramters
    Dim Records()               As String
    Dim dict                    As Dictionary
    Dim Record                  As Variant
    Dim MatchedRecords          As String
    Dim TotalMatchingRows       As Long
    Dim firstRow                As Long
    Dim LastRow                 As Long
    Dim i                       As Long
    Dim tmp                     As Variant
    
    'Get Driving parameters
    firstRow = LBound(ContextTable, 1)
    LastRow = UBound(ContextTable, 1)
    ReDim tmp(0)
    
    'Get matching Rows!
    MatchedRecords = vbNullString
    TotalMatchingRows = 0
    Set dict = New Dictionary
    For i = firstRow + 1 To LastRow
        If Not dict.exists(ContextTable(i, TargetColumn)) Then
            dict.Add ContextTable(i, TargetColumn), ContextTable(i, TargetColumn)
            MatchedRecords = MatchedRecords & i & SEPERATOR_ITEM
            TotalMatchingRows = TotalMatchingRows + 1
        End If
    Next i
    
    'remove last seperator from the matchedrecords variable
    Set dict = Nothing
    If MatchedRecords = vbNullString Then
        GetUniqueListFromTableColumn = tmp
        Exit Function
    Else
        MatchedRecords = VBA.Left$(MatchedRecords, VBA.Len(MatchedRecords) - VBA.Len(SEPERATOR_ITEM))
    End If
    
    'Redimentioning the tmp array
    ReDim tmp(TotalMatchingRows)
    
    'Transferring Data
    Records = VBA.Split(MatchedRecords, SEPERATOR_ITEM)
    i = 1
    
    For Each Record In Records
        tmp(i) = ContextTable(Record, TargetColumn)
        i = i + 1
    Next Record
    
    GetUniqueListFromTableColumn = Me.SortOneDimensionalArrayWithNumbers(tmp)
    
End Function

'@Description("Following function act as Vlookup")
Public Function GetLookupValue(ByVal TargetTable As Variant, _
                                ByVal LookupValue As Variant, _
                                ByVal LookupIndex As Long, _
                                ByVal TargetIndex As Long) As String
Attribute GetLookupValue.VB_Description = "Following function act as Vlookup"
    Dim i   As Long
    Dim LB  As Long
    Dim UB  As Long
    
    LB = LBound(TargetTable, 1)
    UB = UBound(TargetTable, 1)
    
    For i = LB To UB
        If VBA.IsNumeric(LookupValue) Then
            If TargetTable(i, LookupIndex) = LookupValue Then
                GetLookupValue = TargetTable(i, TargetIndex)
                Exit Function
            End If
        Else
            If UCase(TargetTable(i, LookupIndex)) = UCase(LookupValue) Then
                GetLookupValue = TargetTable(i, TargetIndex)
                Exit Function
            End If
        End If
    Next i
    
End Function

'@Description("The Following function will return Column Position from the Headers Array and Column's Name")
Public Function GetTargetColumnIndex(ByVal TargetColumn As String, ByVal Headers As Variant) As Long
Attribute GetTargetColumnIndex.VB_Description = "The Following function will return Column Position from the Headers Array and Column's Name"
    Dim Header  As Variant
    Dim i       As Long
    
    i = 0
    
    For Each Header In Headers
        If Header = TargetColumn Then
            GetTargetColumnIndex = i
            Exit Function
        End If
        i = i + 1
    Next Header
    
End Function

'@Description("Get Target Row of the Table based on Target Index! Like say we are looking for the index no. 4 _
                but it is possible that index/record no. 4 could be found at row now. 3 or 9. Depends on the _
                possibilities.")
Public Function GetTargetRowIndex(ByVal TargetTable As Variant, ByVal TargetIndex As Long, ByVal TargetColumn As Long) As Long
    
    Dim i   As Long
    
    For i = LBound(TargetTable, 1) To UBound(TargetTable, 1)
        If TargetTable(i, TargetColumn) = TargetIndex Then
            GetTargetRowIndex = i
            Exit Function
        End If
    Next i
    
End Function

'@Description("Following function will get us the Maximum number from the Table Column by looking through it")
Public Function GetMax(ByVal TargetTable As Variant, ByVal TargetColumn As Long) As Long
Attribute GetMax.VB_Description = "Following function will get us the Maximum number from the Table Column by looking through it"
    
    Dim i       As Long
    Dim LB      As Long
    Dim UB      As Long
    Dim tmp     As Long
    
    LB = LBound(TargetTable, 1)
    UB = UBound(TargetTable, 1)
    tmp = 0
    
    For i = LB + 1 To UB Step 1
        If TargetTable(i, TargetColumn) > tmp Then
            tmp = TargetTable(i, TargetColumn)
        End If
    Next i
    
    GetMax = tmp

End Function

'@Description("Following function will Sort One Dimensional Array in Ascending Order!")
Public Function SortOneDimensionalArrayWithNumbers(ByVal Arr As Variant) As Variant
Attribute SortOneDimensionalArrayWithNumbers.VB_Description = "Following function will Sort One Dimensional Array in Ascending Order!"
    
    Dim FirstItemIndex  As Long
    Dim LastItemIndex   As Long
    Dim tmp             As Variant
    Dim i               As Long
    Dim j               As Long
    
    FirstItemIndex = LBound(Arr)
    LastItemIndex = UBound(Arr)
    
    For i = FirstItemIndex + 1 To LastItemIndex - 1
        For j = i + 1 To LastItemIndex
            If VBA.IsNumeric(Arr(i)) And VBA.IsNumeric(Arr(j)) Then
                If Val(Arr(i)) > Val(Arr(j)) Then
                    tmp = Arr(j)
                    Arr(j) = Arr(i)
                    Arr(i) = tmp
                End If
            Else
                If Arr(i) > Arr(j) Then
                    tmp = Arr(j)
                    Arr(j) = Arr(i)
                    Arr(i) = tmp
                End If
            End If
        Next j
    Next i
    
    SortOneDimensionalArrayWithNumbers = Arr
    
End Function

'@Description("NumberOfArrayDimensions _
            This function returns the number of dimensions of an array. An unallocated dynamic array _
            has 0 dimensions. This condition can also be tested with IsArrayEmpty. Copy Right @ Chip Pearson")
Private Function NumberOfArrayDimensions(Arr As Variant) As Integer

    Dim Ndx As Integer
    Dim Res As Integer
    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(Arr, Ndx)
    Loop Until Err.Number <> 0
    Err.Clear
    NumberOfArrayDimensions = Ndx - 1
End Function

Private Sub Class_Terminate()
    Set This.ObjRegEx = Nothing
    Disposable.TryDispose This.Algo
End Sub
