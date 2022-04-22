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

Private Function IRepository_GetRecords(ByVal TargetNamedRange As String) As Variant
    IRepository_GetRecords = DataBase.Range(TargetNamedRange)
End Function

Private Sub IRepository_AddRecord(ByVal TargetNamedRange As String, ParamArray LineItems() As Variant)
    'Variable Declaration
    Dim TargetRange     As Range
    Dim Item            As Variant
    Dim TotalItems      As Long
    Dim TRow            As Long
    Dim index           As Long
    Dim I               As Long
    
    'error handelling
    On Error GoTo EhhErr:
    
    'Collect & configure some driving parameters
    TotalItems = UBound(LineItems) + 1
    With DataBase.Range(TargetNamedRange)
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
    For Each Item In LineItems
        TargetRange.Cells(TRow, I).Value = Item
        I = I + 1
    Next Item
    
LeaveThisProcedure:
    
    'Exit condtion
    Set TargetRange = Nothing
    Exit Sub
    
EhhErr:
    MsgBox "Error during Adding the new record" & _
         vbNewLine & vbNewLine & Err.Number & _
         vbNewLine & vbNewLine & Err.Description, vbCritical, "Validator"
    Resume LeaveThisProcedure
End Sub

Private Sub IRepository_DeleteRecord(ByVal TargetNamedRange As String, ByVal TargetIndex As Long)
    'Variable Declaration
    Dim TargetRange     As Range
    Dim TRow            As Long
    Dim R               As Range
    Dim I               As Long
    
    'error handelling
    On Error GoTo EhhErr:
    
    'Collect & configure some driving parameters
    Set TargetRange = DataBase.Range(TargetNamedRange)
    TRow = TargetIndex + 1
    
    'delte the row from the target range
    TargetRange.Rows(TRow).Delete xlShiftUp
    
    're-indexing
    I = 1
    For Each R In TargetRange.Rows
        If VBA.IsNumeric(R.Cells(1, 1).Value) Then
            R.Cells(1, 1).Value = I
            I = I + 1
        End If
    Next R
    
LeaveThisProcedure:
    
    'Exit condtion
    Set TargetRange = Nothing
    Exit Sub
    
EhhErr:
    MsgBox "Error during Deleting the record" & _
         vbNewLine & vbNewLine & Err.Number & _
         vbNewLine & vbNewLine & Err.Description, vbCritical, "Validator"
    Resume LeaveThisProcedure
End Sub

Private Sub IRepository_UpdateRecord(ByVal TargetNamedRange As String, ByVal TargetIndex As Long, ParamArray LineItems() As Variant)
    'Variable Declaration
    Dim TargetRange     As Range
    Dim Item            As Variant
    Dim TotalItems      As Long
    Dim TRow            As Long
    Dim I               As Long
    
    'error handelling
    On Error GoTo EhhErr:
    
    'Collect & configure some driving parameters
    TotalItems = UBound(LineItems) + 1
    Set TargetRange = DataBase.Range(TargetNamedRange)
    TRow = TargetIndex + 1
    
    'init loop paramters
    I = 2
    
    'Insert mechanism
    For Each Item In LineItems
        TargetRange.Cells(TRow, I).Value = Item
        I = I + 1
    Next Item
    
LeaveThisProcedure:
    
    'Exit condtion
    Set TargetRange = Nothing
    Exit Sub
    
EhhErr:
    MsgBox "Error during Updating the record" & _
         vbNewLine & vbNewLine & Err.Number & _
         vbNewLine & vbNewLine & Err.Description, vbCritical, "Validator"
    Resume LeaveThisProcedure
End Sub
