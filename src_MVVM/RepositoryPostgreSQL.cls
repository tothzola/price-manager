VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RepositoryPostgreSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.Repository")
Option Explicit

Implements IRepository

'-------------------------------------------------------------------------
'Private Type Componenets / Enums
'-------------------------------------------------------------------------

Private Type TPostgreSQlRepositoryComponents
    IsEverythingOkay As Boolean
    TargetTableName As String
    TargetTableHeaders As Variant
    ConnectionString As String
    connection As ADODB.connection
    recordset As ADODB.recordset
    SQLQuery As String
End Type

Private This As TPostgreSQlRepositoryComponents

'-------------------------------------------------------------------------
'Public/Private Properties
'-------------------------------------------------------------------------

Private Property Get TargetTableName() As String
    TargetTableName = This.TargetTableName
End Property

Private Property Let TargetTableName(ByVal vNewValue As String)
    This.TargetTableName = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get TargetTableHeaders() As Variant
    TargetTableHeaders = This.TargetTableHeaders
End Property

Private Property Let TargetTableHeaders(ByVal vNewValue As Variant)
    This.TargetTableHeaders = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get IRepository_IsEverythingOkay() As Boolean
    IRepository_IsEverythingOkay = This.IsEverythingOkay
End Property

Private Property Let IRepository_IsEverythingOkay(ByVal RHS As Boolean)
    This.IsEverythingOkay = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get connection() As ADODB.connection
    Set connection = This.connection
End Property

Private Property Set connection(ByVal vNewValue As ADODB.connection)
    Set This.connection = vNewValue
End Property

'-------------------------------------------------------------------------

Private Property Get recordset() As ADODB.recordset
    Set recordset = This.recordset
End Property

Private Property Set recordset(ByVal vNewValue As ADODB.recordset)
    Set This.recordset = vNewValue
End Property

'-------------------------------------------------------------------------
'InIt Method
'-------------------------------------------------------------------------

Private Sub IRepository_InItRepository(ByVal ContextTableName As String, _
                                        Optional ContextConnectionString As String = vbNullString, _
                                        Optional ContextTableHeaders As Variant = vbNullString)
    TargetTableName = ContextTableName
    This.ConnectionString = ContextConnectionString
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
        This.SQLQuery = "SELECT * FROM " & TargetTableName & ";"
        'Executing Query
        recordset.Open This.SQLQuery, connection, adOpenForwardOnly, adLockReadOnly
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
    LogManager.Log ErrorLevel, "Error: " & Err.Number & vbTab & Err.Description
    IRepository_IsEverythingOkay = False
    Resume CleanExit
    
End Function

'-------------------------------------------------------------------------

Private Function IRepository_GetRecordsByCondition(ByVal ColumnIndex As Long, ByVal condition As Variant) As Variant
    
    Call OpenConnection
    
    If IRepository_IsEverythingOkay Then
        On Error GoTo EhhErr
    'Building SQL Qyery String!
        This.SQLQuery = "SELECT * FROM " & TargetTableName & " " & _
                        "WHERE " & TargetTableHeaders(1, ColumnIndex) & " = '" & condition & "';"
        'Executing Query
        recordset.Open This.SQLQuery, connection, adOpenForwardOnly, adLockReadOnly
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
    LogManager.Log ErrorLevel, "Error: " & Err.Number & vbTab & Err.Description
    IRepository_IsEverythingOkay = False
    Resume CleanExit
    
End Function

'-------------------------------------------------------------------------

Private Sub IRepository_AddRecord(ParamArray LineItems() As Variant)
    
    Dim i As Long
    Dim HeadersString As String
    Dim ValuesString As String
    
    Call OpenConnection
    
    If IRepository_IsEverythingOkay Then
        On Error GoTo EhhErr
    'Building SQL Qyery String!
        'Syntax
        'INSERT INTO table_name (column1, column2, column3, ...)
        'VALUES (value1, value2, value3, ...);
        This.SQLQuery = "INSERT INTO " & TargetTableName & " ("
        For i = 2 To UBound(TargetTableHeaders, 2)
            HeadersString = HeadersString & vbNullString & TargetTableHeaders(1, i) & ", "
            If VBA.IsEmpty(LineItems(i - 2)) Then
                ValuesString = ValuesString & "NULL, "
            ElseIf AppMethods.IsCurrency(LineItems(i - 2)) Then
                ValuesString = ValuesString & AppMethods.TransformToDBCompatibleCurrency(LineItems(i - 2)) & ", "
            Else
                ValuesString = ValuesString & "'" & LineItems(i - 2) & "', "
            End If
        Next i
        HeadersString = VBA.Left$(HeadersString, VBA.Len(HeadersString) - 2)
        ValuesString = VBA.Left$(ValuesString, VBA.Len(ValuesString) - 2)
        This.SQLQuery = This.SQLQuery & HeadersString & ") VALUES (" & ValuesString & ");"
        'Executing Query
        recordset.Open This.SQLQuery, connection, adOpenDynamic, adLockOptimistic
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
    LogManager.Log ErrorLevel, "Error: " & Err.Number & vbTab & Err.Description
    IRepository_IsEverythingOkay = False
    Resume CleanExit
    
End Sub

'-------------------------------------------------------------------------

Private Sub IRepository_UpdateRecord(ByVal TargetIndex As Long, ParamArray LineItems() As Variant)
    
    Dim i As Long
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
        This.SQLQuery = "UPDATE " & TargetTableName & " SET "
        For i = 2 To UBound(TargetTableHeaders, 2)
            If VBA.IsEmpty(LineItems(i - 2)) Then
                tmp = tmp & TargetTableHeaders(1, i) & " = NULL, "
            ElseIf AppMethods.IsCurrency(LineItems(i - 2)) Then
                tmp = tmp & TargetTableHeaders(1, i) & " = " & AppMethods.TransformToDBCompatibleCurrency(LineItems(i - 2)) & ", "
            Else
                tmp = tmp & TargetTableHeaders(1, i) & " = '" & LineItems(i - 2) & "', "
            End If
        Next i
        tmp = VBA.Left$(tmp, VBA.Len(tmp) - 2)
        This.SQLQuery = This.SQLQuery & tmp & " WHERE Index = " & TargetIndex & ";"
        'Executing Query
        recordset.Open This.SQLQuery, connection, adOpenDynamic, adLockOptimistic
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
    LogManager.Log ErrorLevel, "Error: " & Err.Number & vbTab & Err.Description
    IRepository_IsEverythingOkay = False
    Resume CleanExit
    
End Sub

'-------------------------------------------------------------------------

Private Sub IRepository_DeleteRecord(ByVal TargetIndex As Long)
    
    Call OpenConnection
    
    If IRepository_IsEverythingOkay Then
        On Error GoTo EhhErr
        'Building SQL Qyery String!
        This.SQLQuery = "DELETE FROM " & TargetTableName & " WHERE Index = " & TargetIndex & ";"
        'Executing Query
        recordset.Open This.SQLQuery, connection, adOpenDynamic, adLockOptimistic
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
    LogManager.Log ErrorLevel, "Error: " & Err.Number & vbTab & Err.Description
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
        .ConnectionString = This.ConnectionString
        .Open
    End With
    
    IRepository_IsEverythingOkay = True
    
    Exit Sub
    
EhhErr:
    MsgBox Err.Description, vbCritical, SIGN
    Err.Clear
    LogManager.Log ErrorLevel, "Error: " & Err.Number & vbTab & Err.Description
    IRepository_IsEverythingOkay = False
    
End Sub

Private Sub CloseConnection()
    If CBool(connection.State And adStateOpen) = True Then connection.Close
    This.SQLQuery = vbNullString
End Sub

Private Sub CloseRecordSet()
    If CBool(recordset.State And adStateOpen) = True Then recordset.Close
End Sub

Private Function GetDataWithHeadings() As Variant
    If Not (recordset.BOF Or recordset.EOF) Then
        Dim tmp As Variant
        tmp = Application.WorksheetFunction.Transpose(recordset.GetRows)
        GetDataWithHeadings = AppMethods.JoinTables(TargetTableHeaders, tmp)
    Else
        GetDataWithHeadings = TargetTableHeaders
    End If
End Function

'-------------------------------------------------------------------------
'Class Events
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    TargetTableName = vbNullString
    TargetTableHeaders = Empty
    This.ConnectionString = vbNullString
    Set connection = Nothing
    Set recordset = Nothing
End Sub
