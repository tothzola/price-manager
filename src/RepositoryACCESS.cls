VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryACCESS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Repository"
Option Explicit

Implements IRepository

'-------------------------------------------------------------------------
'Private Type Componenets / Enums
'-------------------------------------------------------------------------

Private Type TAccessRepositoryComponents
    IsEverythingOkay As Boolean
    TargetTableName As String
    TargetTableHeaders As Variant
    ConnectionString As String
    connection As ADODB.connection
    recordset As ADODB.recordset
    SQLQuery As String
End Type

Private this As TAccessRepositoryComponents

'-------------------------------------------------------------------------
'Public/Private Properties
'-------------------------------------------------------------------------

Private Property Get TargetTableName() As String
    TargetTableName = this.TargetTableName
End Property

Private Property Let TargetTableName(ByVal vNewValue As String)
    this.TargetTableName = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get TargetTableHeaders() As Variant
    TargetTableHeaders = this.TargetTableHeaders
End Property

Private Property Let TargetTableHeaders(ByVal vNewValue As Variant)
    this.TargetTableHeaders = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get IRepository_IsEverythingOkay() As Boolean
    IRepository_IsEverythingOkay = this.IsEverythingOkay
End Property

Private Property Let IRepository_IsEverythingOkay(ByVal RHS As Boolean)
    this.IsEverythingOkay = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get connection() As ADODB.connection
    Set connection = this.connection
End Property

Private Property Set connection(ByVal vNewValue As ADODB.connection)
    Set this.connection = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get recordset() As ADODB.recordset
    Set recordset = this.recordset
End Property

Private Property Set recordset(ByVal vNewValue As ADODB.recordset)
    Set this.recordset = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt Method
'-------------------------------------------------------------------------

Private Sub IRepository_InItRepository(ByVal ContextTableName As String, _
                                        Optional ContextConnectionString As String = vbNullString, _
                                        Optional ContextTableHeaders As Variant = vbNullString)
    TargetTableName = ContextTableName
    this.ConnectionString = ContextConnectionString
    TargetTableHeaders = ContextTableHeaders
    If connection Is Nothing Then Set connection = New ADODB.connection
    If recordset Is Nothing Then Set recordset = New ADODB.recordset
End Sub

'-------------------------------------------------------------------------
'Methods
'-------------------------------------------------------------------------

Private Sub IRepository_CheckTableStatus()
    Call OpenConnection
    Call CloseConnection
End Sub

'-------------------------------------------------------------------------

Private Function IRepository_GetRecords() As Variant
    
    Call OpenConnection
    
    If IRepository_IsEverythingOkay Then
        On Error GoTo EhhErr
    'Building SQL Qyery String!
        this.SQLQuery = "SELECT * FROM " & TargetTableName & ";"
        'Executing Query
        recordset.Open this.SQLQuery, connection, adOpenForwardOnly, adLockReadOnly
        'Return Table!
        IRepository_GetRecords = GetDataWithHeadings
    End If
    
    'if you are here then operation is successful!
    IRepository_IsEverythingOkay = True
    
CleanExit:
    Call CloseRecordSet
    Call CloseConnection
    Exit Function
    
EhhErr:
    If Err Then MsgBox Err.Description, vbCritical, SIGN
    Err.Clear
    IRepository_IsEverythingOkay = False
    Resume CleanExit
    
End Function

'-------------------------------------------------------------------------

Private Function IRepository_GetRecordsByCondition(ByVal columnIndex As Long, ByVal condition As Variant) As Variant
    
    Call OpenConnection
    
    If IRepository_IsEverythingOkay Then
        On Error GoTo EhhErr
    'Building SQL Qyery String!
        If VBA.IsNumeric(condition) Then
            this.SQLQuery = "SELECT * FROM " & TargetTableName & " WHERE [" & TargetTableHeaders(1, columnIndex) & "] = " & condition & ";"
        Else
            this.SQLQuery = "SELECT * FROM " & TargetTableName & " WHERE [" & TargetTableHeaders(1, columnIndex) & "] = '" & condition & "';"
        End If
        'Executing Query
        recordset.Open this.SQLQuery, connection, adOpenForwardOnly, adLockReadOnly
        'Return Table!
        IRepository_GetRecordsByCondition = GetDataWithHeadings
    End If
    
    'if you are here then operation is successful!
    IRepository_IsEverythingOkay = True
    
CleanExit:
    Call CloseRecordSet
    Call CloseConnection
    Exit Function
    
EhhErr:
    If Err Then MsgBox Err.Description, vbCritical, SIGN
    Err.Clear
    IRepository_IsEverythingOkay = False
    Resume CleanExit
    
End Function

'-------------------------------------------------------------------------

Private Sub IRepository_AddRecord(ParamArray LineItems() As Variant)
    
    Dim I As Long
    Dim HeadersString As String
    Dim ValuesString As String
    
    Call OpenConnection
    
    If IRepository_IsEverythingOkay Then
        On Error GoTo EhhErr
    'Building SQL Qyery String!
        'Syntax
        'INSERT INTO table_name ([column1], [column2], [column3], ...)
        'VALUES (value1, value2, value3, ...);
        this.SQLQuery = "INSERT INTO " & TargetTableName & " ("
        For I = 2 To UBound(TargetTableHeaders, 2)
            HeadersString = HeadersString & "[" & TargetTableHeaders(1, I) & "], "
            If VBA.IsEmpty(LineItems(I - 2)) Then
                ValuesString = ValuesString & "'', "
            Else
                ValuesString = ValuesString & "'" & LineItems(I - 2) & "', "
            End If
        Next I
        HeadersString = VBA.Left$(HeadersString, VBA.Len(HeadersString) - 2)
        ValuesString = VBA.Left$(ValuesString, VBA.Len(ValuesString) - 2)
        this.SQLQuery = this.SQLQuery & HeadersString & ") VALUES (" & ValuesString & ");"
        'Executing Query
        recordset.Open this.SQLQuery, connection, adOpenDynamic, adLockOptimistic
    End If
    
    'if you are here then operation is successful!
    IRepository_IsEverythingOkay = True
    
CleanExit:
    Call CloseRecordSet
    Call CloseConnection
    Exit Sub
    
EhhErr:
    If Err Then MsgBox Err.Description, vbCritical, SIGN
    Err.Clear
    IRepository_IsEverythingOkay = False
    Resume CleanExit
    
End Sub

'-------------------------------------------------------------------------

Private Sub IRepository_UpdateRecord(ByVal TargetIndex As Long, ParamArray LineItems() As Variant)
    
    Dim I As Long
    Dim tmp As String
    
    Call OpenConnection

    If IRepository_IsEverythingOkay Then
        On Error GoTo EhhErr
    'Building SQL Qyery String!
        'Syntax
        'Update table_name
        'SET column1 = value1, column2 = value2, ...
        'WHERE columnX = valueX;
        tmp = vbNullString
        this.SQLQuery = "UPDATE " & TargetTableName & " SET "
        For I = 2 To UBound(TargetTableHeaders, 2)
            If VBA.IsEmpty(LineItems(I - 2)) Then
                tmp = tmp & "[" & TargetTableHeaders(1, I) & "] = '', "
            Else
                tmp = tmp & "[" & TargetTableHeaders(1, I) & "] = '" & LineItems(I - 2) & "', "
            End If
        Next I
        tmp = VBA.Left$(tmp, VBA.Len(tmp) - 2)
        this.SQLQuery = this.SQLQuery & tmp & " WHERE [Index] = " & TargetIndex & ";"
        'Executing Query
        recordset.Open this.SQLQuery, connection, adOpenDynamic, adLockOptimistic
    End If
    
    'if you are here then operation is successful!
    IRepository_IsEverythingOkay = True
    
CleanExit:
    Call CloseRecordSet
    Call CloseConnection
    Exit Sub
    
EhhErr:
    If Err Then MsgBox Err.Description, vbCritical, SIGN
    Err.Clear
    IRepository_IsEverythingOkay = False
    Resume CleanExit
    
End Sub

'-------------------------------------------------------------------------

Private Sub IRepository_DeleteRecord(ByVal TargetIndex As Long)
    
    Call OpenConnection
    
    If IRepository_IsEverythingOkay Then
        On Error GoTo EhhErr
        'Building SQL Qyery String!
        this.SQLQuery = "DELETE FROM " & TargetTableName & " WHERE [Index] = " & TargetIndex & ";"
        'Executing Query
        recordset.Open this.SQLQuery, connection, adOpenDynamic, adLockOptimistic
    End If
    
    'if you are here then operation is successful!
    IRepository_IsEverythingOkay = True
    
CleanExit:
    Call CloseRecordSet
    Call CloseConnection
    Exit Sub
    
EhhErr:
    If Err Then MsgBox Err.Description, vbCritical, SIGN
    Err.Clear
    IRepository_IsEverythingOkay = False
    Resume CleanExit
    
End Sub

'-------------------------------------------------------------------------
'Private Methods and Helper Functions
'-------------------------------------------------------------------------

Private Sub OpenConnection()

    On Error GoTo EhhErr
    
    If connection Is Nothing Then Set connection = New ADODB.connection
    
    With connection
        .ConnectionString = this.ConnectionString
        .Open
    End With
    
    IRepository_IsEverythingOkay = True
    
    Exit Sub
    
EhhErr:
    MsgBox Err.Description, vbCritical, SIGN
    Err.Clear
    IRepository_IsEverythingOkay = False
    
End Sub

Private Sub CloseConnection()
    If CBool(connection.State And adStateOpen) = True Then connection.Close
    this.SQLQuery = vbNullString
End Sub

Private Sub CloseRecordSet()
    If CBool(recordset.State And adStateOpen) = True Then recordset.Close
End Sub

Private Function GetDataWithHeadings() As Variant
    If Not (recordset.BOF Or recordset.EOF) Then
        Dim tmp As Variant
        tmp = Application.WorksheetFunction.Transpose(recordset.GetRows)
        GetDataWithHeadings = JoinTables(TargetTableHeaders, tmp)
    Else
        GetDataWithHeadings = TargetTableHeaders
    End If
End Function

Private Function JoinTables(ByVal Table1 As Variant, ByVal Table2 As Variant) As Variant
    Dim TotalRowsTable1     As Long
    Dim TotalRowsTable2     As Long
    Dim TotalColumnsTable1  As Long
    Dim TotalColumnsTable2  As Long
    Dim FinalTableRows      As Long
    Dim FinalTableColumns   As Long
    Dim I                   As Long
    Dim J                   As Long
    Dim FinalTable          As Variant
    'Check if the table2 is Double dimensional or not?
    'if not then create it
    If NumberOfArrayDimensions(Table2) = 1 Then
        Dim tmp As Variant
        ReDim tmp(1 To 1, 1 To UBound(Table1, 2))
        For I = 1 To UBound(Table1, 2)
            tmp(1, I) = Table2(I)
        Next I
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
    For I = 1 To FinalTableRows
        For J = 1 To FinalTableColumns
            If I <= TotalRowsTable1 Then
                FinalTable(I, J) = Table1(I, J)
            Else
                FinalTable(I, J) = Table2(I - TotalRowsTable1, J)
            End If
        Next J
    Next I
    JoinTables = FinalTable
End Function

'Copy Right @ Chip Pearson
Private Function NumberOfArrayDimensions(arr As Variant) As Integer
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' NumberOfArrayDimensions
    ' This function returns the number of dimensions of an array. An unallocated dynamic array
    ' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Ndx As Integer
    Dim Res As Integer
    On Error Resume Next
    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(arr, Ndx)
    Loop Until Err.Number <> 0
    Err.Clear
    NumberOfArrayDimensions = Ndx - 1
End Function

'-------------------------------------------------------------------------
'Class Events
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    TargetTableName = vbNullString
    TargetTableHeaders = Empty
    this.ConnectionString = vbNullString
    Set connection = Nothing
    Set recordset = Nothing
End Sub


