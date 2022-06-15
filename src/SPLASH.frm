VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SPLASH 
   Caption         =   "InitScreen"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7095
   OleObjectBlob   =   "SPLASH.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SPLASH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "View"
Option Explicit

Private ControlView As IFormControl

Private Sub UserForm_Initialize()

    Set ControlView = FormControl.Create
    With ControlView
        .ShowTitleBar UF:=Me, HideTitle:=True
        .SetFormOpacity UF:=Me, Opacity:=200
    End With
    
End Sub
