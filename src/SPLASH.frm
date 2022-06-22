VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SPLASH 
   Caption         =   "InitScreen"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4830
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
Private Const PROGRESSBAR_MAXWIDTH As Integer = 224

Public Event Activated()
Public Event Cancelled()

'Private Sub UserForm_Initialize()
'
'
'
'End Sub

Private Sub UserForm_Activate()
    
    Set ControlView = FormControl.Create
    With ControlView
        .ShowTitleBar UF:=Me, HideTitle:=True
        .SetFormOpacity UF:=Me, Opacity:=200
    End With
    
    ProgressBar.Width = 0                        ' it's set to 10 to be visible at design-time
    RaiseEvent Activated
End Sub

Public Sub Update(ByVal percentValue As Single, Optional ByVal labelValue As String, Optional ByVal captionValue As String)

    If labelValue <> vbNullString Then
        ProgressLabel.Caption = labelValue
    End If

    If captionValue <> vbNullString Then
        Me.Caption = captionValue
    End If

    ProgressBar.Width = percentValue * PROGRESSBAR_MAXWIDTH
    DoEvents

End Sub

Public Property Get CurrentProgress() As Double
    CurrentProgress = ProgressBar.Width / PROGRESSBAR_MAXWIDTH
End Property


