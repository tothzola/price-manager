VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"

'@ModuleAttribute VB_Ext_KEY, "Rubberduck", "Predeclared Class Module"
'@ModuleDescription "An object keeping the AppState in scope."
'@Folder("AppObjects")
'@PredeclaredId
'@Exposed
Option Explicit

Implements IAppState
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

Private this As TState

'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As AppState
    Set Class = AppState
End Property

'@Description "Creates a new application context."
Public Function Create(Optional ByVal ConnectToDB As RepositoryType = RepositoryType.TYPE_POSTGRESQL, _
        Optional ByVal MainService As IServices, _
        Optional ByVal UserService As IServices) As IAppState
    Guard.NonDefaultInstance AppState
    
    Dim result As AppState
    Set result = New AppState
    
    InitServicesMain result, MainService
    InitServicesUser result, UserService
    
    Set Create = result
End Function

Private Sub InitServicesMain(ByVal State As AppState, ByVal Service As IServices)
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

Private Sub InitServicesUser(ByVal State As AppState, ByVal Service As IServices)
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

Private Property Get DBConnectionType() As RepositoryType
    DBConnectionType = this.DBConnectionType
End Property

Private Property Let DBConnectionType(ByVal RHS As RepositoryType)
    this.DBConnectionType = RHS
End Property

'@Description "Gets/sets the Services Main reference."
Public Property Get MainService() As IServices
    Guard.NonDefaultInstance AppState
    Set MainService = this.MainService
End Property

Public Property Set MainService(ByVal RHS As IServices)
    Guard.NonDefaultInstance AppState
    Guard.NullReference RHS
    Set this.MainService = RHS
End Property

'@Description "Gets/sets the Services User reference."
Public Property Get UserService() As IServices
    Guard.NonDefaultInstance AppState
    Set UserService = this.UserService
End Property

Public Property Set UserService(ByVal RHS As IServices)
    Guard.NonDefaultInstance AppState
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

Private Property Get IAppState_MainService() As IServices
    Set IAppState_MainService = MainService
End Property

Private Property Get IAppState_UserService() As IServices
    Set IAppState_UserService = UserService
End Property

Private Function IAppState_IsAppOnline() As Boolean

    Dim result As Boolean
    
    'Check if the machine is connected to the internet
    If InternetGetConnectedState(Stat, 0&) <> 0 Then
        result = (InitMainServicesPassed And InitUserServicesPassed)
    Else
        VBA.MsgBox "Check Internet Connection", vbCritical, Title:=SIGN
    End If
    
    IAppState_IsAppOnline = result
    
End Function

Private Sub IDisposable_Dispose()
    Dispose
End Sub


