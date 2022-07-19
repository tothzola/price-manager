Attribute VB_Name = "modSingletons"
'@Folder("GlobalEntities")
Option Explicit

Private Type TSingletonComponents
    GetMethod As AppMethods
    SendEmail As EmailServices
End Type

Private this As TSingletonComponents

Public Function GetMethod() As AppMethods
    If this.GetMethod Is Nothing Then Set this.GetMethod = New AppMethods
    Set GetMethod = this.GetMethod
End Function

Public Function SendEmail() As EmailServices
    If this.SendEmail Is Nothing Then Set this.SendEmail = New EmailServices
    Set SendEmail = this.SendEmail
End Function
