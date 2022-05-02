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
    IServices_GetListOfAllRecordsFromContextTable = Repository.GetRecords(ContextTableName)
End Function

Private Function IServices_GetSpecificListOfRecordsFromContextTable(ByVal ColumnReference As Long, ByVal SearchCondition As Variant) As Variant
    If VBA.IsNumeric(SearchCondition) Then
        IServices_GetSpecificListOfRecordsFromContextTable = Repository.GetRecordsByCondition(ContextTableName, ColumnReference, Val(SearchCondition))
    Else
        IServices_GetSpecificListOfRecordsFromContextTable = Repository.GetRecordsByCondition(ContextTableName, ColumnReference, SearchCondition)
    End If
End Function

Private Sub IServices_AddNewRecordToContextTable(ParamArray contextLineItems() As Variant)
    Call Repository.AddRecord(ContextTableName, _
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
End Sub

Private Sub IServices_UpdateRecordOfContextTable(ByVal index As Long, ParamArray contextLineItems() As Variant)
    Call Repository.UpdateRecord(ContextTableName, _
                                    index, _
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
End Sub

Private Sub IServices_DeleteRecordofContextTable(ByVal index As Long)
    Call Repository.DeleteRecord(ContextTableName, index)
End Sub

