VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object keeping the AppState in scope."
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@ModuleDescription "An object keeping the AppState in scope."
'@Folder("PriceApproval.Model")
'@PredeclaredId
'@Exposed
Option Explicit

Implements IAppContext
Implements IDisposable

Private Type TState
    DBConnectionType As RepositoryType
    MainService As IServices
    UserService As IServices
    
    Disposed As Boolean
End Type

Private this As TState

'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As AppContext
Attribute Class.VB_Description = "Returns class reference"
    Set Class = AppContext
End Property

'@Description "Creates a new application context."
Public Function Create(Optional ByVal ConnectToDB As RepositoryType = RepositoryType.TYPE_POSTGRESQL, _
                       Optional ByVal MainService As IServices, _
                       Optional ByVal UserService As IServices) As IAppContext
Attribute Create.VB_Description = "Creates a new application context."
    Guard.NonDefaultInstance AppContext
    
    Dim result As AppContext
    Set result = New AppContext
    
        result.DBConnectionType = ConnectToDB
    
        InitServicesMain result, MainService
        InitServicesUser result, UserService
    
    Set Create = result
End Function

Private Sub InitServicesMain(ByVal State As AppContext, ByVal Service As IServices)
    Dim tempService As IServices
    If Service Is Nothing Then
        'default to concrete implementation:
        Dim ConcreteService As ServicesMain
        Set ConcreteService = New ServicesMain
        Set tempService = ConcreteService
    Else
        'use injected instance (could be a test stub):
        Set tempService = Service
    End If
    Set State.MainService = tempService
End Sub

Private Sub InitServicesUser(ByVal State As AppContext, ByVal Service As IServices)
    Dim tempService As IServices
    If Service Is Nothing Then
        'default to concrete implementation:
        Dim ConcreteService As ServicesUser
        Set ConcreteService = New ServicesUser
        Set tempService = ConcreteService
    Else
        'use injected instance (could be a test stub):
        Set tempService = Service
    End If
    Set State.UserService = tempService
End Sub

Private Function InitMainServicesPassed() As Boolean
                            
    MainService.InItService DBConnectionType, _
                            PriceApprovalDataResources.MAIN_TABLE_NAME, _
                            PriceApprovalDataResources.arrListOfColumns_MAIN_Table, _
                            PriceApprovalDataResources.GetConnectionString(DBConnectionType)
    
    InitMainServicesPassed = MainService.IsEverythingOkayInEngine

End Function

Private Function InitUserServicesPassed() As Boolean
                            
    UserService.InItService DBConnectionType, _
                            PriceApprovalDataResources.USERS_TABLE_NAME, _
                            PriceApprovalDataResources.arrListOfColumns_USERS_TABLE, _
                            PriceApprovalDataResources.GetConnectionString(DBConnectionType)

    InitUserServicesPassed = UserService.IsEverythingOkayInEngine

End Function

Public Property Get DBConnectionType() As RepositoryType
    Guard.DefaultInstance Me
    DBConnectionType = this.DBConnectionType
End Property

Friend Property Let DBConnectionType(ByVal RHS As RepositoryType)
    Guard.DefaultInstance Me
    this.DBConnectionType = RHS
End Property

'@Description "Gets/sets the Services Main reference."
Public Property Get MainService() As IServices
Attribute MainService.VB_Description = "Gets/sets the Services Main reference."
    Guard.DefaultInstance Me
    Set MainService = this.MainService
End Property

Friend Property Set MainService(ByVal RHS As IServices)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization this.MainService
    Guard.NullReference RHS
    Set this.MainService = RHS
End Property

'@Description "Gets/sets the Services User reference."
Public Property Get UserService() As IServices
Attribute UserService.VB_Description = "Gets/sets the Services User reference."
    Guard.DefaultInstance Me
    Set UserService = this.UserService
End Property

Friend Property Set UserService(ByVal RHS As IServices)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization this.UserService
    Guard.NullReference RHS
    Set this.UserService = RHS
End Property

Private Sub Dispose()

    If this.Disposed Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If

    If Not this.MainService Is Nothing Then
        Disposable.TryDispose this.MainService
        Set this.MainService = Nothing
    End If
    
    If Not this.UserService Is Nothing Then
        Disposable.TryDispose this.UserService
        Set this.UserService = Nothing
    End If
    
    this.Disposed = True
    
    #If TestMode Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " is terminating"
    #End If
    
End Sub

Private Sub Class_Terminate()
    If Not this.Disposed Then Dispose
End Sub

Private Property Get IAppContext_MainService() As IServices
    Set IAppContext_MainService = MainService
End Property

Private Property Get IAppContext_UserService() As IServices
    Set IAppContext_UserService = UserService
End Property

Private Function IAppContext_IsRepositoryReachable() As Boolean

    Dim result As Boolean
    
    '@Ignore UnassignedVariableUsage
    If System.ConnectedToNetwork Then
        result = (InitMainServicesPassed And InitUserServicesPassed)
    Else
        VBA.MsgBox "Check Internet Connection !", vbCritical, Title:=PriceApprovalSignature
    End If
    
    IAppContext_IsRepositoryReachable = result
    
End Function

Private Sub IDisposable_Dispose()
    Dispose
End Sub
