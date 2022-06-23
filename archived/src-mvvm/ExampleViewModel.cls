VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExampleViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A StringFormat example ViewModel implementation for some dialog."
'@ModuleDescription "A StringFormat example ViewModel implementation for some dialog."
'@Folder("Example")
'@PredeclaredId
'@Exposed
Option Explicit

Private Type TViewModel
    Title As String
    
    FormatToGeneral As String
    ValueForGeneral As String
    
    FormatToCustom As String
    ValueForCustom As String
End Type

Private this As TViewModel

Public Function Create() As ExampleViewModel
    
    Dim result As ExampleViewModel
    Set result = New ExampleViewModel
    
    Set Create = result
End Function

Public Property Get Title() As String
    Title = this.Title
End Property

Public Property Let Title(ByVal RHS As String)
    this.Title = RHS
End Property


Public Property Get FormatToGeneral() As String
    FormatToGeneral = this.FormatToGeneral
End Property

Public Property Let FormatToGeneral(ByVal RHS As String)
    this.FormatToGeneral = RHS
End Property


Public Property Get FormatToCustom() As String
    FormatToCustom = this.FormatToCustom
End Property

Public Property Let FormatToCustom(ByVal RHS As String)
    this.FormatToCustom = RHS
End Property


Public Property Get ValueForGeneral() As String
    ValueForGeneral = this.ValueForGeneral
End Property

Public Property Let ValueForGeneral(ByVal RHS As String)
    If this.ValueForGeneral <> RHS Then
        this.ValueForGeneral = RHS
    End If
End Property


Public Property Get ValueForCustom() As String
    ValueForCustom = this.ValueForCustom
End Property

Public Property Let ValueForCustom(ByVal RHS As String)
    If this.ValueForCustom <> RHS Then
        this.ValueForCustom = RHS
    End If
End Property

