VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ServicesUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("MVVM.Services")
'@PredeclaredId
Option Explicit

Implements IServices
Implements IDisposable

Private Type TUserServicesComponents
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

Private This As TUserServicesComponents

'-------------------------------------------------------------------------

Private Property Get IServices_IsEverythingOkayInEngine() As Boolean
    IServices_IsEverythingOkayInEngine = This.IsEverythingOkayInEngine
End Property

Private Property Let IServices_IsEverythingOkayInEngine(ByVal RHS As Boolean)
    This.IsEverythingOkayInEngine = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get RepositoryTypes() As RepositoryType
    RepositoryTypes = This.RepositoryTypes
End Property

Private Property Let RepositoryTypes(ByVal RHS As RepositoryType)
    This.RepositoryTypes = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get Repository() As IRepository

    Set Repository = This.Repository
End Property

Private Property Set Repository(ByVal RHS As IRepository)

    Set This.Repository = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get ConnectionString() As String
    ConnectionString = This.ConnectionString
End Property

Private Property Let ConnectionString(ByVal RHS As String)
    This.ConnectionString = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get ContextTableName() As String
    ContextTableName = This.ContextTableName
End Property

Private Property Let ContextTableName(ByVal RHS As String)
    This.ContextTableName = RHS
End Property

'-------------------------------------------------------------------------

Private Property Get ContextTableHeaders() As Variant
    Dim tmp As Variant
    Dim i As Long
    ReDim tmp(1 To 1, 1 To 7)
    For i = 1 To UBound(tmp, 2)
        tmp(1, i) = This.ContextTableHeaders(i - 1)
    Next i
    ContextTableHeaders = tmp
End Property

Private Property Let ContextTableHeaders(ByVal RHS As Variant)
    This.ContextTableHeaders = RHS
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

    Dim cryptPass As String
    cryptPass = AppMethods.Algo.ComputeHash(ContextLineItems(4))
    ContextLineItems(4) = cryptPass

    Call Repository.AddRecord(ContextLineItems(0), _
                              ContextLineItems(1), _
                              ContextLineItems(2), _
                              ContextLineItems(3), _
                              ContextLineItems(4), _
                              ContextLineItems(5))
    IServices_IsEverythingOkayInEngine = Repository.IsEverythingOkay
    
End Sub

Private Sub IServices_UpdateRecordOfContextTable(ByVal Index As Long, ParamArray ContextLineItems() As Variant)
    
    Dim cryptPass As String
    cryptPass = AppMethods.Algo.ComputeHash(ContextLineItems(4))
    ContextLineItems(4) = cryptPass
        
    Call Repository.UpdateRecord(Index, _
                                 ContextLineItems(0), _
                                 ContextLineItems(1), _
                                 ContextLineItems(2), _
                                 ContextLineItems(3), _
                                 ContextLineItems(4), _
                                 ContextLineItems(5))
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
    If Not This.Disposed Then Dispose
End Sub

Private Sub Dispose()

    If This.Disposed Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If

    If Not This.Repository Is Nothing Then
        Disposable.TryDispose This.Repository
        Set This.Repository = Nothing
    End If
    
    This.Disposed = True
    
    #If TestMode Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " is terminating"
    #End If
    
End Sub

Private Sub IDisposable_Dispose()
    Dispose
End Sub

