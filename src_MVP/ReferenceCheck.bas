Attribute VB_Name = "ReferenceCheck"
Attribute VB_Description = "Internal methods used to configure this addin on startup."
'@Folder("System.Settings")
'@ModuleDescription("Internal methods used to configure this addin on startup.")
Option Explicit
Option Private Module

Public Const APP_NAME As String = "PriceApproval_MVP"
Public Const APP_ADDIN_NAME As String = APP_NAME & ".xlam"

Public Const DEVELOPER_ZOLTAN As String = "Toth, Zoltan"
Public Const DEVELOPER_KAMAL As String = "Bharakhda, Kamal"

Public Const DEVELOPER_NAME As String = "Developer Name"
Public Const DEVELOPER_EMAIL As String = "developer@Email.com"

'@EntryPoint
Public Function CheckReferenceCompatibility() As Boolean

    On Error GoTo CleanFail
    
    Dim result As Boolean
    
    If Not IsValidApplicationFileName(ThisWorkbook.Name, APP_ADDIN_NAME) Then
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
                                                 CommonDllVbProjectReference.RefEdit + _
                                                 CommonDllVbProjectReference.Outlook) _
        Then
        
        ReferenceResolver.DisplayReferenceError DEVELOPER_NAME, DEVELOPER_EMAIL
        ExitApp
    Else
        result = True
    End If

CleanExit:
    CheckReferenceCompatibility = result
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
               vbCritical, SIGN & " - Error: Unauthorized Name Change"

    End If
    
    IsValidApplicationFileName = result
    
End Function

Public Sub ExitApp()

    Dim isDevelopers As Boolean
    isDevelopers = (Application.UserName = DEVELOPER_ZOLTAN) Or (Application.UserName = DEVELOPER_KAMAL)
    
    If isDevelopers Then
        MsgBox "Application is opened by Developers." & vbNewLine & "App Autoclose cancelled.", vbInformation, SIGN
    Else
        ThisWorkbook.Saved = True
        ThisWorkbook.Close
    End If
    
End Sub

Private Sub ManageApplicationStartupError()
    MsgBox "Application StartUp Error" & VBA.Constants.vbNewLine & VBA.Constants.vbNewLine & _
           "An error occured while this application was attempting to load. " & _
           "If this issue persists, please contact the developer of this project. " & VBA.Constants.vbNewLine & VBA.Constants.vbNewLine & _
           "This application will now exit.", vbCritical, SIGN
           
    ExitApp
End Sub

