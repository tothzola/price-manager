VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServicesUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Services"
Option Explicit

Implements IServices

Private Type TUserServicesComponents
    'Public Properties
    IsEverythingOkayInEngine As Boolean
    'Private properties
    Repository As IRepository
    RepositoryTypes As RepositoryType
    ContextTableName As String
    ConnectionString As String
    ContextTableHeaders As Variant
End Type

Private this As TUserServicesComponents

'-------------------------------------------------------------------------

Private Property Get IServices_IsEverythingOkayInEngine() As Boolean
    IServices_IsEverythingOkayInEngine = this.IsEverythingOkayInEngine
End Property

Private Property Let IServices_IsEverythingOkayInEngine(ByVal RHS As Boolean)
    this.IsEverythingOkayInEngine = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get RepositoryTypes() As RepositoryType
    RepositoryTypes = this.RepositoryTypes
End Property

Private Property Let RepositoryTypes(ByVal RHS As RepositoryType)
    this.RepositoryTypes = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get Repository() As IRepository
    Set Repository = this.Repository
End Property

Private Property Set Repository(ByVal RHS As IRepository)
    Set this.Repository = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get ConnectionString() As String
    ConnectionString = this.ConnectionString
End Property

Private Property Let ConnectionString(ByVal RHS As String)
    this.ConnectionString = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get ContextTableName() As String
    ContextTableName = this.ContextTableName
End Property

Private Property Let ContextTableName(ByVal RHS As String)
    this.ContextTableName = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get ContextTableHeaders() As Variant
    Dim tmp As Variant
    Dim I As Long
    ReDim tmp(1 To 1, 1 To 6)
    For I = 1 To UBound(tmp, 2)
        tmp(1, I) = this.ContextTableHeaders(I - 1)
    Next I
    ContextTableHeaders = tmp
End Property

Private Property Let ContextTableHeaders(ByVal RHS As Variant)
    this.ContextTableHeaders = RHS
End Property

'-------------------------------------------------------------------------
'Contstructor
'-------------------------------------------------------------------------

Private Sub IServices_InItService(ByVal SelectedRepositoryType As RepositoryType, _
                                    ByVal TableName As String, _
                                    ByVal TableHeaders As Variant, _
                                    Optional ByVal ContextConnectionString As String = vbNullString)
    'Hydrate Properties
    RepositoryTypes = SelectedRepositoryType
    ConnectionString = ContextConnectionString
    ContextTableName = TableName
    ContextTableHeaders = TableHeaders
    
    Select Case RepositoryTypes
        Case RepositoryType.TYPE_EXCEL_NAMED_RANGE
            If Repository Is Nothing Then Set Repository = New RepositoryExcel
            Call Repository.InItRepository(ContextTableName)
        Case RepositoryType.TYPE_SHAREPOINT_LIST
            If Repository Is Nothing Then Set Repository = New RepositorySharePoint
            Call Repository.InItRepository(ContextTableName, ConnectionString, ContextTableHeaders)
        Case RepositoryType.TYPE_MYSQL
            If Repository Is Nothing Then Set Repository = New RepositoryMySQL
            Call Repository.InItRepository(ContextTableName, ConnectionString, ContextTableHeaders)
        Case RepositoryType.TYPE_ACCESS
            If Repository Is Nothing Then Set Repository = New RepositoryACCESS
            Call Repository.InItRepository(ContextTableName, ConnectionString, ContextTableHeaders)
    End Select
    
    'Check for the Table Status and hence we can indirectly get the status of Database itself
    Call Repository.CheckTableStatus
    IServices_IsEverythingOkayInEngine = Repository.IsEverythingOkay
    
End Sub

'-------------------------------------------------------------------------
'Methods
'-------------------------------------------------------------------------

Private Function IServices_GetListOfAllRecordsFromContextTable() As Variant
    IServices_GetListOfAllRecordsFromContextTable = Repository.GetRecords()
    IServices_IsEverythingOkayInEngine = Repository.IsEverythingOkay
End Function

Private Function IServices_GetSpecificListOfRecordsFromContextTable(ByVal ColumnReference As Long, ByVal SearchCondition As Variant) As Variant
    IServices_GetSpecificListOfRecordsFromContextTable = Repository.GetRecordsByCondition(ColumnReference, SearchCondition)
    IServices_IsEverythingOkayInEngine = Repository.IsEverythingOkay
End Function

Private Sub IServices_AddNewRecordToContextTable(ParamArray ContextLineItems() As Variant)
    Call Repository.AddRecord(ContextLineItems(0), _
                              ContextLineItems(1), _
                              ContextLineItems(2), _
                              ContextLineItems(3), _
                              ContextLineItems(4))
    IServices_IsEverythingOkayInEngine = Repository.IsEverythingOkay
End Sub

Private Sub IServices_UpdateRecordOfContextTable(ByVal index As Long, ParamArray ContextLineItems() As Variant)
    Call Repository.UpdateRecord(index, _
                                 ContextLineItems(0), _
                                 ContextLineItems(1), _
                                 ContextLineItems(2), _
                                 ContextLineItems(3), _
                                 ContextLineItems(4))
    IServices_IsEverythingOkayInEngine = Repository.IsEverythingOkay
End Sub

Private Sub IServices_DeleteRecordofContextTable(ByVal index As Long)
    Call Repository.DeleteRecord(index)
    IServices_IsEverythingOkayInEngine = Repository.IsEverythingOkay
End Sub

'-------------------------------------------------------------------------
'Class Events
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    Set Repository = Nothing
End Sub

