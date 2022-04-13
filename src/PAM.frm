VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PAM 
   Caption         =   "Price Approval Manager V1.0"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2535
   OleObjectBlob   =   "PAM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'SETTINGS
Const MESSAGE_WELCOMESCREEN_LOGOUT_STATE As String = "Welcome to The Price Approval Manager"
Const MESSAGE_WELCOMESCREEN_LOGIN_STATE As String = "Welcome "

'Private Variables/Objects
Private ExtendedMethods As MultiFrameViewExtended

Private Sub cmdApproverLogout_Click()
    'Logout State
    Call ExtendedMethods.ActivateFrames(Me.frameLogin, Me.frameWelcome)
    Call UpdateWelcomeMessageForLogoutState
End Sub

Private Sub cmdCancelExportUtility_Click()
    'cancel
    Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameWelcome)
    Call UpdateWelcomeMessage
End Sub

Private Sub cmdCancelFromLoginInterface_Click()
    'Logout State
    Call ExtendedMethods.ActivateFrames(Me.frameLogin, Me.frameWelcome)
    Call UpdateWelcomeMessageForLogoutState
End Sub

Private Sub cmdCancelPriceFormInterface_Click()
    'back to the dashboard
    Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.frameWelcome)
    Call UpdateWelcomeMessage
End Sub

Private Sub cmdCancelRecordContainer_Click()
    'back to the dashboard
    With ExtendedMethods
        If Me.lblActiveUserType.Caption = "Client" Then
            Call .ActivateFrames(Me.frameClient, Me.frameWelcome)
        Else
            Call .ActivateFrames(Me.frameApprover, Me.frameWelcome)
        End If
    End With
    Call UpdateWelcomeMessage
End Sub

Private Sub cmdCancelUserManager_Click()
    'cancel
    Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameWelcome)
    Call UpdateWelcomeMessage
End Sub

Private Sub cmdClientLogout_Click()
    'Logout State
    Call ExtendedMethods.ActivateFrames(Me.frameLogin, Me.frameWelcome)
    Call UpdateWelcomeMessageForLogoutState
End Sub

Private Sub cmdExit_Click()
    'get exit from the application
    Me.Hide
End Sub

Private Sub cmdLogin_Click()
    'login
    If Me.txtUsername_VT1 = "Kamal" And Me.txtPassword_VT0 = "123" Then
        Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.frameWelcome)
        Call UpdateActiveUserInfomation("Kamal", "Client", "Active")
        Call UpdateWelcomeMessage
    ElseIf Me.txtUsername_VT1 = "Zoltan" And Me.txtPassword_VT0 = "123" Then
        Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameWelcome)
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
    Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameRecordsContainer)
    'hide Approve Rejection Buttons
    ShowApprovalRejectionButtons True
End Sub

Private Sub cmdOpenClientHistory_Click()
    'open client history interface
    Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.frameRecordsContainer)
    'hide Approve Rejection Buttons
    ShowApprovalRejectionButtons False
End Sub

Private Sub cmdOpenExportUtility_Click()
    'open export utility
    Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameExportUtility)
End Sub

Private Sub cmdOpenLoginInterface_Click()
    'open login interface
    Call ExtendedMethods.ActivateFrames(Me.frameLogin, Me.frameLoginInterface)
    'set focus
    Me.txtUsername_VT1.SetFocus
End Sub

Private Sub cmdOpenPendingList_Click()
    'open pending list for approver
    Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameRecordsContainer)
    'hide Approve Rejection Buttons
    ShowApprovalRejectionButtons True
End Sub

Private Sub cmdOpenPriceForm_Click()
    'open Price Form Interface
    Call ExtendedMethods.ActivateFrames(Me.frameClient, Me.framePriceForm)
End Sub

Private Sub cmdOpenUserManager_Click()
    'open user manager
    Call ExtendedMethods.ActivateFrames(Me.frameApprover, Me.frameUserManager)
End Sub

Private Sub UserForm_Initialize()
    'InItApp
    Call InItApplication
End Sub

'Private methods

Private Sub InItApplication()
    'init Extended Methods
    Set ExtendedMethods = New MultiFrameViewExtended
    With ExtendedMethods
        'Re-Dimension UserForm
        Set .TargetForm = Me
        .formWidth = 600
        .formHeight = 360
        Call .ReDimensionForm
        'Always On Frame Properties
        Set .frameAlwaysOn = Me.frameInfo
        .alwaysOnTop = 6
        .alwaysOnLeft = 6
        .alwaysOnWidth = 140
        .alwaysOnHeight = 78
        'Side Panel Frames Properties
        .sideFrameTop = 90
        .sideFrameLeft = 6
        .sideFrameWidth = 140
        .sideFrameHeight = 234
        'Main Panel Frames Properties
        .mainFrameTop = 6
        .mainFrameLeft = 152
        .mainFrameWidth = 430
        .mainFrameHeight = 318
        'InIt Interface
        Call .ActivateFrames(Me.frameLogin, Me.frameWelcome)
        Call UpdateWelcomeMessageForLogoutState
    End With
End Sub

Private Sub ShowApprovalRejectionButtons(ByVal Decision As Boolean)
    With Me
        .cmdApproveRecord.Visible = Decision
        .cmdRejectRecord.Visible = Decision
    End With
End Sub

Private Sub UpdateWelcomeMessage()
    Call ExtendedMethods.ChangeControlProperties(Me.lblWelcomeMessage, MESSAGE_WELCOMESCREEN_LOGIN_STATE & Me.lblActiveUsername.Caption)
End Sub

Private Sub UpdateWelcomeMessageForLogoutState()
    With ExtendedMethods
        Call .ChangeControlProperties(Me.lblWelcomeMessage, MESSAGE_WELCOMESCREEN_LOGOUT_STATE)
        Call .SetStateofControlsToNullState(Me.lblActiveUsername, Me.lblActiveUserType, Me.lblActiveUserStatus)
    End With
End Sub

Private Sub UpdateActiveUserInfomation(ByVal uName As String, ByVal uType As String, ByVal uStatus As String)
    With ExtendedMethods
        Call .ChangeControlProperties(Me.lblActiveUsername, uName)
        Call .ChangeControlProperties(Me.lblActiveUserType, uType)
        If uStatus = "Active" Then
            Call .ChangeControlProperties(Me.lblActiveUserStatus, uStatus, &H8000&)
        Else
            Call .ChangeControlProperties(Me.lblActiveUserStatus, uStatus, vbRed)
        End If
    End With
End Sub



