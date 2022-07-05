VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServicesMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("MVVM.Services")
'@PredeclaredId
Option Explicit

Implements IServices
Implements IDisposable

Private Type TMainServicesComponents
    'Public Properties
    IsEverythingOkayInEngine As Boolean
    'Private properties
    Repository As IRepository
    RepositoryTypes As RepositoryType
    ContextTableName As String
    ConnectionString As String
    ContextTableHeaders As Variant
    Disposed As Boolean
End Type

Private this As TMainServicesComponents

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
    Dim i As Long
    ReDim tmp(1 To 1, 1 To 16)
    For i = 1 To UBound(tmp, 2)
        tmp(1, i) = this.ContextTableHeaders(i - 1)
    Next i
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
    ContextTableName = TableName
    ContextTableHeaders = TableHeaders
    ConnectionString = ContextConnectionString
    
    Select Case RepositoryTypes
    Case RepositoryType.TYPE_POSTGRESQL
        If Repository Is Nothing Then Set Repository = New RepositoryPostgreSQL
        Repository.InItRepository ContextTableName, _
                                  ConnectionString, _
                                  ContextTableHeaders
            
    Case RepositoryType.TYPE_ACCESS
        If Repository Is Nothing Then Set Repository = New RepositoryACCESS
        Repository.InItRepository ContextTableName, _
                                  ConnectionString, _
                                  ContextTableHeaders
            
    End Select
    
    'Check for the Table Status and hence we can indirectly get the status of Database itself
    Repository.CheckTableStatus
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
                              ContextLineItems(4), _
                              ContextLineItems(5), _
                              ContextLineItems(6), _
                              ContextLineItems(7), _
                              ContextLineItems(8), _
                              ContextLineItems(9), _
                              ContextLineItems(10), _
                              ContextLineItems(11), _
                              ContextLineItems(12), _
                              ContextLineItems(13), _
                              ContextLineItems(14))
    IServices_IsEverythingOkayInEngine = Repository.IsEverythingOkay
End Sub

Private Sub IServices_UpdateRecordOfContextTable(ByVal Index As Long, ParamArray ContextLineItems() As Variant)
    Call Repository.UpdateRecord(Index, _
                                 ContextLineItems(0), _
                                 ContextLineItems(1), _
                                 ContextLineItems(2), _
                                 ContextLineItems(3), _
                                 ContextLineItems(4), _
                                 ContextLineItems(5), _
                                 ContextLineItems(6), _
                                 ContextLineItems(7), _
                                 ContextLineItems(8), _
                                 ContextLineItems(9), _
                                 ContextLineItems(10), _
                                 ContextLineItems(11), _
                                 ContextLineItems(12), _
                                 ContextLineItems(13), _
                                 ContextLineItems(14))
    IServices_IsEverythingOkayInEngine = Repository.IsEverythingOkay
End Sub

Private Sub IServices_DeleteRecordofContextTable(ByVal Index As Long)
    Call Repository.DeleteRecord(Index)
    IServices_IsEverythingOkayInEngine = Repository.IsEverythingOkay
End Sub

'-------------------------------------------------------------------------
'Class Events
'-------------------------------------------------------------------------

Private Sub Class_Terminate()
    If Not this.Disposed Then Dispose
End Sub

Private Sub Dispose()

    If this.Disposed Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If

    If Not this.Repository Is Nothing Then
        Disposable.TryDispose this.Repository
        Set this.Repository = Nothing
    End If

    
    this.Disposed = True
    
    #If TestMode Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " is terminating"
    #End If
    
End Sub

Private Sub IDisposable_Dispose()
    Dispose
End Sub

