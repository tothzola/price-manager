Attribute VB_Name = "ReferenceCheck"
Attribute VB_Description = "Internal methods used to configure this addin on startup."
'@Folder("System.Reference")
'@ModuleDescription("Internal methods used to configure this addin on startup.")
Option Explicit

'*******************************************************************************
'Allows developer to save if the shift is pressed
#If VBA7 Then
    Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal vKey As Long) As Long
#Else
    Private Declare Function GetKeyState Lib "user32" (ByVal vKey As Long) As Long
#End If

Public DoSave As Boolean

Public Const DEVELOPER_KAMAL As String = "Bharakhda, Kamal"

Public Const DEVELOPER_NAME As String = "Toth, Zoltan"
Public Const DEVELOPER_EMAIL As String = "Zoltan.Toth@freudenberg-pm.com"

Private Const KEY_MASK As Long = &HFF80
Private Const SHIFT_KEY As Long = &H10

'@EntryPoint
Public Function SystemCompatibility(ByVal AppName As String) As Boolean

    On Error GoTo CleanFail
    
    Dim result As Boolean
    
    If Not IsValidApplicationFileName(ThisWorkbook.Name, AppName) Then
        ExitApp
    Else
        result = True
    End If
    
    If Not ReferenceResolver.TryAddDllReferences(dllReference:= _
                                                 CommonDllVbProjectReference.AdoDbRef + _
                                                 CommonDllVbProjectReference.AdoDDlExtRef + _
                                                 CommonDllVbProjectReference.ScriptingRuntimeRef + _
                                                 CommonDllVbProjectReference.VbScriptRegExpRef + _
                                                 CommonDllVbProjectReference.MSForms + _
                                                 CommonDllVbProjectReference.MSOffice + _
                                                 CommonDllVbProjectReference.MSXML + _
                                                 CommonDllVbProjectReference.RefEdit + _
                                                 CommonDllVbProjectReference.Outlook) _
        Then
        
        ReferenceResolver.DisplayReferenceError DEVELOPER_NAME, DEVELOPER_EMAIL
        ExitApp
    Else
        result = True
    End If

CleanExit:
    SystemCompatibility = result
    Exit Function

CleanFail:
    ManageApplicationStartupError
    Resume CleanExit
    Resume
    
End Function

Public Function IsValidApplicationFileName(ByVal currentApplicationFileName As String, ByVal expectedApplicationFileName As String) As Boolean

    Dim result As Boolean
    
    If (currentApplicationFileName = expectedApplicationFileName) Then
        result = True
        
    Else
        MsgBox "Looks like this application's original name has been changed. " & _
               VBA.Constants.vbNewLine & VBA.Constants.vbNewLine & _
               "To use this application, its name must remain as the following:" & _
               VBA.Constants.vbNewLine & VBA.Constants.vbNewLine & _
               expectedApplicationFileName & VBA.Constants.vbNewLine & VBA.Constants.vbNewLine & _
               "Clicking 'Okay' will automatically exit this application. " & _
               "Once closed, you MUST restore the name " & _
               "to the one mentioned above.", _
               vbCritical, PriceApprovalSignature & " - Error: Unauthorized Name Change"

    End If
    
    IsValidApplicationFileName = result
    
End Function

Public Function AllowWorkbookSave(Optional ByVal warningMessage As String = "In order to preserve the structure of this application, saving has been disabled.") As Boolean
    
    If DoSave Then AllowWorkbookSave = DoSave: Exit Function
    
    Dim result As Boolean
    'only allows user to save if they are holding down shift AND they are the admin
    If CBool(GetKeyState(SHIFT_KEY) And KEY_MASK) Then
        If IsApplicationDeveloper(DEVELOPER_NAME) Then
            result = True
        Else
            MsgBox warningMessage, vbExclamation + vbOKOnly, "Requires admin privileges"
            result = False
        End If
    Else
        MsgBox warningMessage, vbExclamation + vbOKOnly, "Requires admin privileges"
        result = False
    End If
    AllowWorkbookSave = result
    
End Function

Public Function IsApplicationDeveloper(ByVal expectedDeveloperUserName As String) As Boolean
    IsApplicationDeveloper = (VBA.UCase$(VBA.Trim$(DEVELOPER_NAME)) = VBA.UCase$(VBA.Trim$(expectedDeveloperUserName)))
End Function

Public Sub ExitApp()

    Dim isDevelopers As Boolean
    isDevelopers = (Application.UserName = DEVELOPER_NAME) Or (Application.UserName = DEVELOPER_KAMAL)
    
    If isDevelopers Then
        MsgBox "Application is opened by Developers." & vbNewLine & "App Autoclose cancelled.", vbInformation, APP_SIGNATURE
    Else
        ThisWorkbook.Saved = True
        ThisWorkbook.Close
    End If
    
End Sub

Private Sub ManageApplicationStartupError()
    MsgBox "Application StartUp Error" & VBA.Constants.vbNewLine & VBA.Constants.vbNewLine & _
           "An error occured while this application was attempting to load. " & _
           "If this issue persists, please contact the developer of this project. " & VBA.Constants.vbNewLine & VBA.Constants.vbNewLine & _
           "This application will now exit.", vbCritical, APP_SIGNATURE
           
    ExitApp
End Sub

