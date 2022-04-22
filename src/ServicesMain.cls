VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServicesMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IServices

Private Type TUserServicesComponents
    Repository As IRepository
    RepositoryType As RepositoryType
    ContextTableName As String
    contextLineItems As Variant
End Type

Private this As TUserServicesComponents

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

Private Property Get contextLineItems() As Variant
    contextLineItems = this.contextLineItems
End Property

Private Property Let IServices_ContextLineItems(ByVal RHS As Variant)
    this.contextLineItems = RHS
End Property

'-------------------------------------------------------------------------
'Contstructor
'-------------------------------------------------------------------------

Private Sub IServices_InItService(ByVal SelectedRepositoryType As RepositoryType, ByVal TableName As String)
    this.RepositoryType = SelectedRepositoryType
    IServices_ContextTableName = TableName
    Select Case this.RepositoryType
        Case RepositoryType.TYPE_EXCEL_NAMED_RANGE
            Set IServices_Repository = New RepositoryExcel
        Case RepositoryType.TYPE_SHAREPOINT_LIST
            Set IServices_Repository = New RepositorySharePoint
        Case RepositoryType.TYPE_MYSQL
            Set IServices_Repository = New RepositoryMySQL
    End Select
End Sub

'-------------------------------------------------------------------------
'Methods
'-------------------------------------------------------------------------

Private Function IServices_GetListOfAllRecordsFromContextTable() As Variant
    Dim tmp As Variant
    tmp = Repository.GetRecords(ContextTableName)
    IServices_GetListOfAllRecordsFromContextTable = tmp
End Function

Private Function IServices_GetSpecificListOfRecordsFromContextTable(ByVal ColumnReference As String, ByVal SearchCondition As String) As Variant
    Debug.Print "Not Implemented Yet"
End Function

Private Sub IServices_AddNewRecordToContextTable(ParamArray contextLineItems() As Variant)
    Debug.Print "Not Implemented Yet"
End Sub

Private Sub IServices_UpdateRecordOfContextTable(ByVal index As Long, ParamArray contextLineItems() As Variant)
    Debug.Print "Not Implemented Yet"
End Sub

Private Sub IServices_DeleteRecordofContextTable(ByVal index As Long)
    Debug.Print "Not Implemented Yet"
End Sub
