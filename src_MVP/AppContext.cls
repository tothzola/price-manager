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

'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@ModuleDescription "An object keeping the AppState in scope."
'@Folder("AppObjects.Model")
'@PredeclaredId
'@Exposed
Option Explicit

Implements IAppContext
Implements IDisposable

#If VBA7 And Win64 Then
    Private Stat As LongPtr
    Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" _
    (lpdwFlags As LongPtr, ByVal dwReserved As Long) As Boolean
#Else
    Private Stat As Long
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" _
                             (lpdwFlags As Long, ByVal dwReserved As Long) As Boolean
#End If

Private Type TState
    DBConnectionType As RepositoryType
    MainService As IServices
    UserService As IServices
    
    Disposed As Boolean
End Type

Private This As TState

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
                            DataResources.MAIN_TABLE_NAME, _
                            DataResources.arrListOfColumns_MAIN_Table, _
                            DataResources.GetConnectionString(DBConnectionType)
    
    InitMainServicesPassed = MainService.IsEverythingOkayInEngine

End Function

Private Function InitUserServicesPassed() As Boolean
                            
    UserService.InItService DBConnectionType, _
                            DataResources.USERS_TABLE_NAME, _
                            DataResources.arrListOfColumns_USERS_TABLE, _
                            DataResources.GetConnectionString(DBConnectionType)

    InitUserServicesPassed = UserService.IsEverythingOkayInEngine

End Function

Public Property Get DBConnectionType() As RepositoryType
    Guard.DefaultInstance Me
    DBConnectionType = This.DBConnectionType
End Property

Public Property Let DBConnectionType(ByVal RHS As RepositoryType)
    Guard.DefaultInstance Me
    This.DBConnectionType = RHS
End Property

'@Description "Gets/sets the Services Main reference."
Public Property Get MainService() As IServices
Attribute MainService.VB_Description = "Gets/sets the Services Main reference."
    Guard.DefaultInstance Me
    Set MainService = This.MainService
End Property

Public Property Set MainService(ByVal RHS As IServices)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization This.MainService
    Guard.NullReference RHS
    Set This.MainService = RHS
End Property

'@Description "Gets/sets the Services User reference."
Public Property Get UserService() As IServices
Attribute UserService.VB_Description = "Gets/sets the Services User reference."
    Guard.DefaultInstance Me
    Set UserService = This.UserService
End Property

Public Property Set UserService(ByVal RHS As IServices)
    Guard.DefaultInstance Me
    Guard.DoubleInitialization This.UserService
    Guard.NullReference RHS
    Set This.UserService = RHS
End Property

Private Sub Dispose()

    If This.Disposed Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If

    If Not This.MainService Is Nothing Then
        Disposable.TryDispose This.MainService
        Set This.MainService = Nothing
    End If
    
    If Not This.UserService Is Nothing Then
        Disposable.TryDispose This.UserService
        Set This.UserService = Nothing
    End If
    
    This.Disposed = True
    
    #If TestMode Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " is terminating"
    #End If
    
End Sub

Private Sub Class_Terminate()
    If Not This.Disposed Then Dispose
End Sub

Private Property Get IAppContext_MainService() As IServices
    Set IAppContext_MainService = MainService
End Property

Private Property Get IAppContext_UserService() As IServices
    Set IAppContext_UserService = UserService
End Property

Private Function IAppContext_IsRepositoryReachable() As Boolean

    Dim result As Boolean
    
    'Check if the machine is connected to the internet
    '@Ignore UnassignedVariableUsage
    If InternetGetConnectedState(Stat, 0&) <> 0 Then
        result = (InitMainServicesPassed And InitUserServicesPassed)
    Else
        VBA.MsgBox "Check Internet Connection !", vbCritical, Title:=SIGN
    End If
    
    IAppContext_IsRepositoryReachable = result
    
End Function

Private Sub IDisposable_Dispose()
    Dispose
End Sub
