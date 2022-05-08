VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServicesMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Services"
Option Explicit

Implements IServices

Private Type TUserServicesComponents
    IsEverythingOkayInEngine As Boolean
    Repository As IRepository
    RepositoryType As RepositoryType
    ContextTableName As String
    ContextTableHeaders As Variant
    contextLineItems As Variant
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

Private Property Get Repository() As IRepository
    Set Repository = this.Repository
End Property

Private Property Set IServices_Repository(ByVal RHS As IRepository)
    Set this.Repository = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get ContextTableName() As String
    ContextTableName = this.ContextTableName
End Property

Private Property Let IServices_ContextTableName(ByVal RHS As String)
    this.ContextTableName = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get ContextTableHeaders() As Variant
    Dim tmp As Variant
    ReDim tmp(1 To 1, 1 To 13)
    tmp(1, 1) = this.ContextTableHeaders(0)
    tmp(1, 2) = this.ContextTableHeaders(1)
    tmp(1, 3) = this.ContextTableHeaders(2)
    tmp(1, 4) = this.ContextTableHeaders(3)
    tmp(1, 5) = this.ContextTableHeaders(4)
    tmp(1, 6) = this.ContextTableHeaders(5)
    tmp(1, 7) = this.ContextTableHeaders(6)
    tmp(1, 8) = this.ContextTableHeaders(7)
    tmp(1, 9) = this.ContextTableHeaders(8)
    tmp(1, 10) = this.ContextTableHeaders(9)
    tmp(1, 11) = this.ContextTableHeaders(10)
    tmp(1, 12) = this.ContextTableHeaders(11)
    tmp(1, 13) = this.ContextTableHeaders(12)
    ContextTableHeaders = tmp
End Property

Private Property Let IServices_ContextTableHeaders(ByVal RHS As Variant)
    this.ContextTableHeaders = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get contextLineItems() As Variant
    contextLineItems = this.contextLineItems
End Property

Private Property Let IServices_ContextLineItems(ByVal RHS As Variant)
    this.contextLineItems = RHS
End Property

'-------------------------------------------------------------------------
'Contstructor
'-------------------------------------------------------------------------

Private Sub IServices_InItService(ByVal SelectedRepositoryType As RepositoryType, ByVal TableName As String, ByVal TableHeaders As Variant)
    this.RepositoryType = SelectedRepositoryType
    IServices_ContextTableName = TableName
    IServices_ContextTableHeaders = TableHeaders
    Select Case this.RepositoryType
        Case RepositoryType.TYPE_EXCEL_NAMED_RANGE
            If Repository Is Nothing Then Set IServices_Repository = New RepositoryExcel
            Call Repository.InItRepository(ContextTableName)
        Case RepositoryType.TYPE_SHAREPOINT_LIST
            If Repository Is Nothing Then Set IServices_Repository = New RepositorySharePoint
            Call Repository.InItRepository(ContextTableName, "Connection String Not Inserted Yet!", ContextTableHeaders)
        Case RepositoryType.TYPE_MYSQL
            If Repository Is Nothing Then Set IServices_Repository = New RepositoryMySQL
            Call Repository.InItRepository(ContextTableName, "Connection String Not Inserted Yet!", ContextTableHeaders)
        Case RepositoryType.TYPE_ACCESS
            If Repository Is Nothing Then Set IServices_Repository = New RepositoryACCESS
            Call Repository.InItRepository(ContextTableName, modDataSources.ConnectionString_ACCESS, ContextTableHeaders)
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

Private Sub IServices_AddNewRecordToContextTable(ParamArray contextLineItems() As Variant)
    Call Repository.AddRecord(contextLineItems(0), _
                              contextLineItems(1), _
                              contextLineItems(2), _
                              contextLineItems(3), _
                              contextLineItems(4), _
                              contextLineItems(5), _
                              contextLineItems(6), _
                              contextLineItems(7), _
                              contextLineItems(8), _
                              contextLineItems(9), _
                              contextLineItems(10), _
                              contextLineItems(11))
    IServices_IsEverythingOkayInEngine = Repository.IsEverythingOkay
End Sub

Private Sub IServices_UpdateRecordOfContextTable(ByVal index As Long, ParamArray contextLineItems() As Variant)
    Call Repository.UpdateRecord(index, _
                                 contextLineItems(0), _
                                 contextLineItems(1), _
                                 contextLineItems(2), _
                                 contextLineItems(3), _
                                 contextLineItems(4), _
                                 contextLineItems(5), _
                                 contextLineItems(6), _
                                 contextLineItems(7), _
                                 contextLineItems(8), _
                                 contextLineItems(9), _
                                 contextLineItems(10), _
                                 contextLineItems(11))
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
    Set IServices_Repository = Nothing
End Sub

