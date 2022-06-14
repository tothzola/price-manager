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

Private Function IRepository_GetRecordsByCondition(ByVal ColumnIndex As Long, ByVal condition As Variant) As Variant
    Dim TargetRange     As Range
    Dim TargetTable     As Variant
    Set TargetRange = DataBase.Range(TargetTableName)
    TargetTable = TargetRange.value
    If VBA.IsNumeric(condition) Then
        IRepository_GetRecordsByCondition = GetMethod.FilterListWithSingleCondition(TargetTable, ColumnIndex, Val(condition))
    Else
        IRepository_GetRecordsByCondition = GetMethod.FilterListWithSingleCondition(TargetTable, ColumnIndex, condition)
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
    Dim i               As Long
    
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
    TargetRange.Cells(TRow, 1).value = index
    
    'init loop paramters
    i = 2
    
    'Insert mechanism
    For Each item In LineItems
        TargetRange.Cells(TRow, i).value = item
        i = i + 1
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
    Dim i               As Long
    
    'error handelling
    On Error GoTo EhhErr:
    
    'Collect & configure some driving parameters
    Set TargetRange = DataBase.Range(TargetTableName)
    TRow = TargetIndex + 1
    
    'delte the row from the target range
    TargetRange.Rows(TRow).Delete xlShiftUp
    
    're-indexing
    i = 1
    For Each r In TargetRange.Rows
        If VBA.IsNumeric(r.Cells(1, 1).value) Then
            r.Cells(1, 1).value = i
            i = i + 1
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
    Dim i               As Long
    
    'error handelling
    On Error GoTo EhhErr:
    
    'Collect & configure some driving parameters
    TotalItems = UBound(LineItems) + 1
    Set TargetRange = DataBase.Range(TargetTableName)
    TRow = TargetIndex + 1
    
    'init loop paramters
    i = 2
    
    'Insert mechanism
    For Each item In LineItems
        TargetRange.Cells(TRow, i).value = item
        i = i + 1
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

