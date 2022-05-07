VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Repository"
Option Explicit

Implements IRepository

'-------------------------------------------------------------------------
'Private Type Componenets
'-------------------------------------------------------------------------

Private Type TExcelRepositoryComponents
    IsEverythingOkay As Boolean
    TargetTableName As String
End Type

Private this As TExcelRepositoryComponents

'-------------------------------------------------------------------------
'Properties
'-------------------------------------------------------------------------

Private Property Get IRepository_IsEverythingOkay() As Boolean
    IRepository_IsEverythingOkay = this.IsEverythingOkay
End Property

Private Property Let IRepository_IsEverythingOkay(ByVal RHS As Boolean)
    this.IsEverythingOkay = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get TargetTableName() As String
    TargetTableName = this.TargetTableName
End Property

Private Property Let TargetTableName(ByVal vNewValue As String)
    this.TargetTableName = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt Method
'-------------------------------------------------------------------------

Private Sub IRepository_InItRepository(ByVal ContextTableName As String, Optional ContextConnectionString As String = vbNullString, Optional ContextTableHeaders As Variant = vbNullString)
    TargetTableName = ContextTableName
End Sub

'-------------------------------------------------------------------------
'Methods
'-------------------------------------------------------------------------

Private Sub IRepository_CheckTableStatus()
    
    On Error GoTo EhhErr
    
    IRepository_IsEverythingOkay = (DataBase.Range(TargetTableName).Address <> vbNullString)
    
CleanExit:
    Exit Sub
    
EhhErr:
    MsgBox Err.Description, vbCritical, SIGN
    Err.Clear
    IRepository_IsEverythingOkay = False
    Resume CleanExit
    
End Sub

'-------------------------------------------------------------------------

Private Function IRepository_GetRecords() As Variant
    
    On Error GoTo EhhErr
    
    IRepository_GetRecords = DataBase.Range(TargetTableName)
    
    'if you are here then operation is successful!
    IRepository_IsEverythingOkay = True
    
CleanExit:
    Exit Function
EhhErr:
    MsgBox Err.Description, vbCritical, SIGN
    Err.Clear
    IRepository_IsEverythingOkay = False
    Resume CleanExit
    
End Function

'-------------------------------------------------------------------------

Private Function IRepository_GetRecordsByCondition(ByVal columnIndex As Long, ByVal condition As Variant) As Variant
    Dim TargetRange     As Range
    Dim TargetTable     As Variant
    Set TargetRange = DataBase.Range(TargetTableName)
    TargetTable = TargetRange.Value
    If VBA.IsNumeric(condition) Then
        IRepository_GetRecordsByCondition = FilterListWithSingleCondition(TargetTable, columnIndex, Val(condition))
    Else
        IRepository_GetRecordsByCondition = FilterListWithSingleCondition(TargetTable, columnIndex, condition)
    End If
    'if you are here then operation is successful!
    IRepository_IsEverythingOkay = True
    
CleanExit:
    Set TargetRange = Nothing
    Exit Function
EhhErr:
    MsgBox Err.Description, vbCritical, SIGN
    Err.Clear
    IRepository_IsEverythingOkay = False
    Resume CleanExit
End Function

'-------------------------------------------------------------------------

Private Sub IRepository_AddRecord(ParamArray LineItems() As Variant)
    'Variable Declaration
    Dim TargetRange     As Range
    Dim item            As Variant
    Dim TotalItems      As Long
    Dim TRow            As Long
    Dim index           As Long
    Dim I               As Long
    
    'error handelling
    On Error GoTo EhhErr:
    
    'Collect & configure some driving parameters
    TotalItems = UBound(LineItems) + 1
    With DataBase.Range(TargetTableName)
        Set TargetRange = .Resize(.Rows.Count + 1, .Columns.Count)
    End With
    With TargetRange
        TRow = .Cells(.Rows.Count, 1).Row
    End With
    
    'get index
    index = TRow - 1
    
    'place index
    TargetRange.Cells(TRow, 1).Value = index
    
    'init loop paramters
    I = 2
    
    'Insert mechanism
    For Each item In LineItems
        TargetRange.Cells(TRow, I).Value = item
        I = I + 1
    Next item
    
    'if you are here then operation is successful!
    IRepository_IsEverythingOkay = True
    
CleanExit:
    
    'Exit condtion
    Set TargetRange = Nothing
    Exit Sub
    
EhhErr:
    MsgBox "Error during Adding the new record" & _
         vbNewLine & vbNewLine & Err.Number & _
         vbNewLine & vbNewLine & Err.Description, vbCritical, "Validator"
    Err.Clear
    IRepository_IsEverythingOkay = False
    Resume CleanExit
End Sub

'-------------------------------------------------------------------------

Private Sub IRepository_DeleteRecord(ByVal TargetIndex As Long)
    'Variable Declaration
    Dim TargetRange     As Range
    Dim TRow            As Long
    Dim r               As Range
    Dim I               As Long
    
    'error handelling
    On Error GoTo EhhErr:
    
    'Collect & configure some driving parameters
    Set TargetRange = DataBase.Range(TargetTableName)
    TRow = TargetIndex + 1
    
    'delte the row from the target range
    TargetRange.Rows(TRow).Delete xlShiftUp
    
    're-indexing
    I = 1
    For Each r In TargetRange.Rows
        If VBA.IsNumeric(r.Cells(1, 1).Value) Then
            r.Cells(1, 1).Value = I
            I = I + 1
        End If
    Next r
    
    'if you are here then operation is successful!
    IRepository_IsEverythingOkay = True
    
CleanExit:
    
    'Exit condtion
    Set TargetRange = Nothing
    Exit Sub
    
EhhErr:
    MsgBox "Error during Deleting the record" & _
         vbNewLine & vbNewLine & Err.Number & _
         vbNewLine & vbNewLine & Err.Description, vbCritical, "Validator"
    Err.Clear
    IRepository_IsEverythingOkay = False
    Resume CleanExit
End Sub

'-------------------------------------------------------------------------

Private Sub IRepository_UpdateRecord(ByVal TargetIndex As Long, ParamArray LineItems() As Variant)
    'Variable Declaration
    Dim TargetRange     As Range
    Dim item            As Variant
    Dim TotalItems      As Long
    Dim TRow            As Long
    Dim I               As Long
    
    'error handelling
    On Error GoTo EhhErr:
    
    'Collect & configure some driving parameters
    TotalItems = UBound(LineItems) + 1
    Set TargetRange = DataBase.Range(TargetTableName)
    TRow = TargetIndex + 1
    
    'init loop paramters
    I = 2
    
    'Insert mechanism
    For Each item In LineItems
        TargetRange.Cells(TRow, I).Value = item
        I = I + 1
    Next item
    
    'if you are here then operation is successful!
    IRepository_IsEverythingOkay = True
    
CleanExit:
    
    'Exit condtion
    Set TargetRange = Nothing
    Exit Sub
    
EhhErr:
    MsgBox "Error during Updating the record" & _
         vbNewLine & vbNewLine & Err.Number & _
         vbNewLine & vbNewLine & Err.Description, vbCritical, "Validator"
    Err.Clear
    IRepository_IsEverythingOkay = False
    Resume CleanExit
    
End Sub

'Private helping function

Public Function FilterListWithSingleCondition(ByVal ContextTable As Variant, ByVal columnIndex As Long, ByVal condition As Variant) As Variant
    'Driving Paramters
    Const CONSIDERHEADERINDEX   As String = "1" & SEPERATOR_ITEM
    Dim Records()               As String
    Dim Record                  As Variant
    Dim tmp                     As Variant
    Dim MatchedRecords          As String
    Dim TotalMatchingRows       As Long
    Dim firstRow                As Long
    Dim lastRow                 As Long
    Dim FirstCol                As Long
    Dim LastCol                 As Long
    Dim I                       As Long
    Dim J                       As Long
    'Get Driving parameters
    firstRow = LBound(ContextTable, 1)
    lastRow = UBound(ContextTable, 1)
    FirstCol = LBound(ContextTable, 2)
    LastCol = UBound(ContextTable, 2)
    ReDim tmp(1 To 1, 1 To LastCol)
    'Get matching Rows!
    MatchedRecords = CONSIDERHEADERINDEX
    TotalMatchingRows = 0
    For I = firstRow To lastRow Step 1
        If ContextTable(I, columnIndex) = condition Then
            MatchedRecords = MatchedRecords & I & SEPERATOR_ITEM
            TotalMatchingRows = TotalMatchingRows + 1
        End If
    Next I
    'remove last seperator from the matchedrecords variable
    If MatchedRecords = CONSIDERHEADERINDEX Then
        For J = FirstCol To LastCol Step 1
            tmp(1, J) = ContextTable(1, J)
        Next J
        FilterListWithSingleCondition = tmp
        Exit Function
    Else
        MatchedRecords = VBA.Left(MatchedRecords, VBA.Len(MatchedRecords) - VBA.Len(SEPERATOR_ITEM))
    End If
    'Redimentioning the tmp array
    ReDim tmp(firstRow To TotalMatchingRows + 1, FirstCol To LastCol)
    'Transferring Data
    Records = VBA.Split(MatchedRecords, SEPERATOR_ITEM)
    J = 1
    For Each Record In Records
        For I = FirstCol To LastCol Step 1
            tmp(J, I) = ContextTable(Record, I)
        Next I
        J = J + 1
    Next Record
    'return value
    FilterListWithSingleCondition = tmp
End Function
