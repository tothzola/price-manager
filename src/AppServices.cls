VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IAppServices

'type

Private Type TAppServicesComponents
    Repository As IAppRepository
    MainTableName As String
    UsersTableName As String
End Type

Private this As TAppServicesComponents

'properties

Private Property Get Repository() As IAppRepository
    Set Repository = this.Repository
End Property

Private Property Set Repository(ByVal vNewValue As IAppRepository)
    Set this.Repository = vNewValue
End Property

Private Property Get MainTableName() As String
    MainTableName = this.MainTableName
End Property

Private Property Let IAppServices_MainTableName(ByVal RHS As String)
    this.MainTableName = RHS
End Property

Private Property Get UsersTableName() As String
    UsersTableName = this.UsersTableName
End Property

Private Property Let IAppServices_UsersTableName(ByVal RHS As String)
    this.UsersTableName = RHS
End Property

'methods for communication with Users Table

Private Sub IAppServices_GetAllUsersList()
    Debug.Print "Not Implemented Yet"
End Sub

Private Sub IAppServices_AddNewUser()
    Debug.Print "Not Implemented Yet"
End Sub

Private Sub IAppServices_UpdateUser()
    Debug.Print "Not Implemented Yet"
End Sub

Private Sub IAppServices_DeleteUser()
    Debug.Print "Not Implemented Yet"
End Sub

'Class Events

Private Sub Class_Initialize()
    Set Repository = New AppRepository
End Sub

Private Sub Class_Terminate()
    Set Repository = Nothing
End Sub

