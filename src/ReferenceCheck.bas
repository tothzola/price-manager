Attribute VB_Name = "ReferenceCheck"
Attribute VB_Description = "Internal methods used to configure this addin on startup."
'@Folder("Settings")
'@ModuleDescription("Internal methods used to configure this addin on startup.")
Option Explicit
Option Private Module

Public Const APP_NAME As String = "PriceApprovalManager"
Public Const APP_ADDIN_NAME As String = APP_NAME & ".xlam"

'@EntryPoint
Public Function CheckReferenceCompatibility() As Boolean

    Dim result As Boolean
    
    On Error GoTo CleanFail
    
    If Not IsValidApplicationFileName(ThisWorkbook.Name, APP_ADDIN_NAME) And _
        Not ReferenceResolver.TryAddDllReferences(dllReference:= _
                                                 CommonDllVbProjectReference.AdoDbRef + _
                                                 CommonDllVbProjectReference.AdoDDlExtRef + _
                                                 CommonDllVbProjectReference.ScriptingRuntimeRef + _
                                                 CommonDllVbProjectReference.VbScriptRegExpRef + _
                                                 CommonDllVbProjectReference.MSForms + _
                                                 CommonDllVbProjectReference.MSOffice + _
                                                 CommonDllVbProjectReference.RefEdit + _
                                                 CommonDllVbProjectReference.Outlook) _
        Then
        
        ReferenceResolver.DisplayReferenceError "Developer", "Developer E-mail"
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
               vbCritical, "Error: Unauthorized Name Change"

    End If
    
    IsValidApplicationFileName = result
    
End Function

Public Sub ExitApp()

    Dim isDevelopers As Boolean
    If Application.userName = "Zoltan, Toth" Or "Bharakhda, Kamal" Then
        isDevelopers = True
        MsgBox "Application is in development mode closeing app cancelled."
    End If
    
    If Not isDevelopers Then
        If Application.Workbooks.Count = 1 Then
            Application.EnableEvents = False     'this will reset itself when Excel is opened again
            Application.Quit
        Else
            ThisWorkbook.Saved = True
            ThisWorkbook.Close
        End If
    End If
    
End Sub

Private Sub ManageApplicationStartupError()
    MsgBox "Application StartUp Error" & VBA.Constants.vbNewLine & VBA.Constants.vbNewLine & _
           "An error occured while this application was attempting to load. " & _
           "If this issue persists, please contact the developer of this project. " & VBA.Constants.vbNewLine & VBA.Constants.vbNewLine & _
           "This application will now exit.", vbCritical, VBA.Constants.vbNullString
           
    ExitApp
End Sub

