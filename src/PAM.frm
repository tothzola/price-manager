VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PAM 
   Caption         =   "Price Approval Manager V1.0"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "PAM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApproverLogout_Click()
    'Logout State
    ActivateFrames Me.frameLogin, Me.frameWelcome
    Call UpdateWelcomeMessageForLogoutState
End Sub

Private Sub cmdCancelExportUtility_Click()
    'cancel
    ActivateFrames Me.frameApprover, Me.frameWelcome
    Call UpdateWelcomeMessage
End Sub

Private Sub cmdCancelFromLoginInterface_Click()
    'Logout State
    ActivateFrames Me.frameLogin, Me.frameWelcome
    Call UpdateWelcomeMessageForLogoutState
End Sub

Private Sub cmdCancelPriceFormInterface_Click()
    'back to the dashboard
    ActivateFrames Me.frameClient, Me.frameWelcome
    Call UpdateWelcomeMessage
End Sub

Private Sub cmdCancelRecordContainer_Click()
    'back to the dashboard
    If Me.lblActiveUserType.Caption = "Client" Then
        ActivateFrames Me.frameClient, Me.frameWelcome
    Else
        ActivateFrames Me.frameApprover, Me.frameWelcome
    End If
    Call UpdateWelcomeMessage
End Sub

Private Sub cmdCancelUserManager_Click()
    'cancel
    ActivateFrames Me.frameApprover, Me.frameWelcome
    Call UpdateWelcomeMessage
End Sub

Private Sub cmdClientLogout_Click()
    'Logout State
    ActivateFrames Me.frameLogin, Me.frameWelcome
    UpdateWelcomeMessageForLogoutState
End Sub

Private Sub cmdExit_Click()
    'get exit from the application
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    'login
    If Me.txtUsername_VT1 = "Kamal" And Me.txtPassword_VT0 = "123" Then
        ActivateFrames Me.frameClient, Me.frameWelcome
        Call UpdateActiveUserInfomation("Kamal", "Client", "Active")
        Call UpdateWelcomeMessage
    ElseIf Me.txtUsername_VT1 = "Zoltan" And Me.txtPassword_VT0 = "123" Then
        ActivateFrames Me.frameApprover, Me.frameWelcome
        Call UpdateActiveUserInfomation("Zoltan", "Approver", "Active")
        Call UpdateWelcomeMessage
    Else
        MsgBox "Wrong Username or Password.", vbCritical, "Validator"
        Me.txtUsername_VT1.Value = vbNullString
        Me.txtPassword_VT0.Value = vbNullString
        Me.txtUsername_VT1.SetFocus
    End If
End Sub

Private Sub cmdOpenAllHistory_Click()
    'open client history interface
    ActivateFrames Me.frameApprover, Me.frameRecordsContainer
    'hide Approve Rejection Buttons
    ShowApprovalRejectionButtons True
End Sub

Private Sub cmdOpenClientHistory_Click()
    'open client history interface
    ActivateFrames Me.frameClient, Me.frameRecordsContainer
    'hide Approve Rejection Buttons
    ShowApprovalRejectionButtons False
End Sub

Private Sub cmdOpenExportUtility_Click()
    'open export utility
    ActivateFrames Me.frameApprover, Me.frameExportUtility
End Sub

Private Sub cmdOpenLoginInterface_Click()
    'open login interface
    ActivateFrames Me.frameLogin, Me.frameLoginInterface
    'set focus
    Me.txtUsername_VT1.SetFocus
End Sub

Private Sub cmdOpenPendingList_Click()
    'open pending list for approver
    ActivateFrames Me.frameApprover, Me.frameRecordsContainer
    'hide Approve Rejection Buttons
    ShowApprovalRejectionButtons True
End Sub

Private Sub cmdOpenPriceForm_Click()
    'open Price Form Interface
    ActivateFrames Me.frameClient, Me.framePriceForm
End Sub

Private Sub cmdOpenUserManager_Click()
    'open user manager
    ActivateFrames Me.frameApprover, Me.frameUserManager
End Sub

Private Sub UserForm_Initialize()
    'InItApp
    InItApplication
End Sub

'Private methods

Private Sub InItApplication()
    'Give Dimention to Userform
    Me.Width = 600
    Me.Height = 360
    'Logout State
    ActivateFrames Me.frameLogin, Me.frameWelcome
    Call UpdateWelcomeMessageForLogoutState
End Sub

Private Sub ShowApprovalRejectionButtons(ByVal Decision As Boolean)
    With Me
        .cmdApproveRecord.Visible = Decision
        .cmdRejectRecord.Visible = Decision
    End With
End Sub

Private Sub UpdateWelcomeMessage()
    Me.lblWelcomeMessage.Caption = "Welcome " & Me.lblActiveUsername.Caption
End Sub

Private Sub UpdateWelcomeMessageForLogoutState()
    With Me
        .lblWelcomeMessage.Caption = "Welcome to the Price Approver Manger..."
        .lblActiveUsername.Caption = vbNullString
        .lblActiveUserType.Caption = vbNullString
        .lblActiveUserStatus.Caption = vbNullString
    End With
End Sub

Private Sub UpdateActiveUserInfomation(ByVal uName As String, ByVal uType As String, ByVal uStatus As String)
    With Me
        .lblActiveUsername.Caption = uName
        .lblActiveUserType.Caption = uType
        .lblActiveUserStatus.Caption = uStatus
        If uStatus = "Active" Then
            .lblActiveUserStatus.ForeColor = &H8000&     'green
        Else
            .lblActiveUserStatus.ForeColor = vbRed
        End If
    End With
End Sub

Private Sub ActivateFrames(ByVal sidePanelFrame As MSForms.Frame, ByVal mainPanelFrame As MSForms.Frame)
    Dim ctrl As MSForms.Control
    For Each ctrl In Me.Controls
        If VBA.TypeName(ctrl) = "Frame" Then
            If ctrl.name = sidePanelFrame.name Then
                Call RedimensioningOfFramesBasedOnNature(sidePanelFrame, "SIDE")
                ctrl.Visible = True
            ElseIf ctrl.name = mainPanelFrame.name Then
                Call RedimensioningOfFramesBasedOnNature(mainPanelFrame, "MAIN")
                ctrl.Visible = True
            ElseIf ctrl.name = Me.frameInfo.name Then 'Always ON
                Call RedimensioningOfFramesBasedOnNature(Me.frameInfo, "INFO")
                ctrl.Visible = True
            Else
                ctrl.Visible = False
            End If
        End If
    Next ctrl
    Set ctrl = Nothing
End Sub

Private Sub RedimensioningOfFramesBasedOnNature(ByVal ctrl As MSForms.Control, ByVal nature As String)
    If nature = "SIDE" Then
        RedimensionTheFrame ctrl, 90, 6, 140, 234
    ElseIf nature = "MAIN" Then
        RedimensionTheFrame ctrl, 6, 152, 430, 318
    Else
        RedimensionTheFrame ctrl, 6, 6, 140, 78
    End If
End Sub

Private Sub RedimensionTheFrame(ByVal ctrl As MSForms.Control, ByVal Top As Integer, ByVal Left As Integer, ByVal Width As Integer, ByVal Height As Integer)
    With ctrl
        .Top = Top
        .Left = Left
        .Width = Width
        .Height = Height
    End With
End Sub




