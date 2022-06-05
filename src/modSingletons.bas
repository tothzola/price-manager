Attribute VB_Name = "modSingletons"
'@Folder("GlobalEntities")
Option Explicit

Private Type TSingletonComponents
    GetMethod As AppMethods
End Type

Private this As TSingletonComponents

Public Function GetMethod() As AppMethods
    If this.GetMethod Is Nothing Then Set this.GetMethod = New AppMethods
    Set GetMethod = this.GetMethod
End Function
