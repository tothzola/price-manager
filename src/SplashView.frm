VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SplashView 
   ClientHeight    =   3630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4785
   OleObjectBlob   =   "SplashView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SplashView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'@Folder("PriceApproval.View")

Option Explicit

Private Const PROGRESSBAR_MAXWIDTH As Integer = 224

Public Event Activated()
Public Event Cancelled()

Private Sub UserForm_Activate()

    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    With FormControl
        .ShowTitleBar outForm:=Me, HideTitle:=True
        .SetFormOpacity outForm:=Me, Opacity:=180
    End With

    LabelDesigner.Caption = VBA.Chr$(169) & " " & DEVELOPER_NAME
    
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


