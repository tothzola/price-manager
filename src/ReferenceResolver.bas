Attribute VB_Name = "ReferenceResolver"
Attribute VB_Description = "Internal methods used to resolve commonly used references in a VBproject."
'@Folder("System.Reference")
'@ModuleDescription("Internal methods used to resolve commonly used references in a VBproject.")
Option Explicit
Option Private Module

Public Enum CommonDllVbProjectReference
    AdoDbRef = 2 ^ 1
    AdoDDlExtRef = 2 ^ 2
    ScriptingRuntimeRef = 2 ^ 3
    VbScriptRegExpRef = 2 ^ 4
    VbaExtensibilityRef = 2 ^ 5
    MSForms = 2 ^ 6
    MSOffice = 2 ^ 7
    MSComctlLib = 2 ^ 8
    MSXML = 2 ^ 9
    RefEdit = 2 ^ 10
    Outlook = 2 ^ 11
End Enum

Public Const REFERENCE_EXISTS_ERROR_NUMBER As Long = 32813

Private Const DEBUG_REFERENCE_SEPARATOR As String = "--------------------------------------------------------------------------" & VBA.Constants.vbNewLine

Public Function TryAddDllReferences(ByVal dllReference As CommonDllVbProjectReference) As Boolean

    Dim result As Boolean
    
    If Not IsTrustVbaSet Then Exit Function
    
    'OLE Automation
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.VbScriptRegExpRef) _
        Then
        result = TryAddDllReference("{00020430-0000-0000-C000-000000000046}}", majorVersion:=2, minorVersion:=0)
    End If
    
    'Microsoft Office 16.0 Object Library
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.MSOffice) _
        Then
        result = TryAddDllReference("{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}", majorVersion:=2, minorVersion:=8)
    End If

    'Microsoft Forms 2.0 Object Library
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.MSForms) _
        Then
        result = TryAddDllReference("{0D452EE1-E08F-101A-852E-02608C4D0BB4}", majorVersion:=2, minorVersion:=0)
    End If
    
    'Microsoft ActiveX Data Objects 6.1 Library
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.AdoDbRef) _
        Then
        result = TryAddDllReference("{B691E011-1797-432E-907A-4D8C69339129}", majorVersion:=6, minorVersion:=1)
    End If
    
    'Microsoft Windows Common Controls 6.0 (SP6)
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.MSComctlLib) _
        Then
        result = TryAddDllReference("{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}", majorVersion:=2, minorVersion:=2)
    End If
    
    'Microsoft VBScript Regular Expressions 5.5
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.VbScriptRegExpRef) _
        Then
        result = TryAddDllReference("{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", majorVersion:=5, minorVersion:=5)
    End If
    
    'Microsoft Scripting Runtime
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.ScriptingRuntimeRef) _
        Then
        result = TryAddDllReference("{420B2830-E718-11CF-893D-00A0C9054228}", majorVersion:=1, minorVersion:=0)
    End If
    
    'Microsoft Visual Basic for Applications Extensibility 5.3
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.VbaExtensibilityRef) _
        Then
        result = TryAddDllReference("{0002E157-0000-0000-C000-000000000046}", majorVersion:=5, minorVersion:=3)
    End If
    
    'ADOX (Microsoft ADO Ext. 6.0 for DDL and Security)
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.AdoDDlExtRef) _
        Then
        result = TryAddDllReference("{00000600-0000-0010-8000-00AA006D2EA4}", majorVersion:=6, minorVersion:=0)
    End If
    
    'RefEdit (Ref Edit Control)
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.RefEdit) _
        Then
        result = TryAddDllReference("{00024517-0000-0000-C000-000000000046}", majorVersion:=1, minorVersion:=2)
    End If
    
    'MSXML2 (Microsoft XML, v6.0)
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.MSXML) _
        Then
        result = TryAddDllReference("{F5078F18-C551-11D3-89B9-0000F81FE221}", majorVersion:=6, minorVersion:=0)
    End If
    
    'Outlook (Microsoft Outlook 16.0 Object Library)
    If EnumHasFlag(dllReference, CommonDllVbProjectReference.Outlook) _
        Then
        result = TryAddDllReference("{00062FFF-0000-0000-C000-000000000046}", majorVersion:=9, minorVersion:=6)
    End If

    TryAddDllReferences = result
End Function

Public Function TryAddDllReference(ByVal GUID As String, ByVal majorVersion As Long, ByVal minorVersion As Long) As Boolean
    On Error Resume Next
    ThisWorkbook.VBProject.References.AddFromGuid GUID, majorVersion, minorVersion
    TryAddDllReference = ((VBA.Information.Err.Number = REFERENCE_EXISTS_ERROR_NUMBER) Or (VBA.Information.Err.Number = 0))
    On Error GoTo 0

End Function

Public Sub DisplayReferenceError(ByVal developerName As String, ByVal developerEmailAddress As String)
    MsgBox "An error occured while attempting to add External reference(s) required " & _
           "for this Application or Trust access to the VBA project object model is not set." & VBA.Constants.vbNewLine & VBA.Constants.vbNewLine & _
           "Please contact " & developerName & " at " & developerEmailAddress, vbCritical, APP_SIGNATURE
End Sub

Private Function EnumHasFlag(ByVal flagsOrDefault As Long, ByVal searchFlag As Long) As Boolean
    EnumHasFlag = ((flagsOrDefault And searchFlag) = searchFlag)
End Function

Private Function IsTrustVbaSet() As Boolean

    Dim abort As Boolean
    Do Until CheckIfTrusted(abort)
        If abort Then Exit Do
    Loop
    IsTrustVbaSet = Not abort
    
End Function

Private Function CheckIfTrusted(ByRef abort As Boolean) As Boolean
    
    On Error Resume Next
    Dim result As Boolean
    result = (ThisWorkbook.VBProject.VBE.VBProjects.Count) > 0
    On Error GoTo 0
    
    If Not result Then

        Dim userSelection As Long
        userSelection = MsgBox( _
                    Prompt:="Programmatic access to Visual Basic Project is not trusted." & VBA.Constants.vbNewLine & VBA.Constants.vbNewLine & _
                             "To resolve, Click ""Retry""" & VBA.Constants.vbNewLine & _
                             "Macro Settings -> Select: ""Enable all macros""" & VBA.Constants.vbNewLine & _
                             "Developer Macro Settings -> Check: ""Trust access to the VBA project object model""", _
                             Buttons:=vbOKCancel + vbDefaultButton1 + vbCritical, Title:=APP_SIGNATURE & " - Trust VBA is not set!")

        If userSelection = 1 Then                    'Ok
            Application.CommandBars.ExecuteMso "MacroSecurity"
            
        End If
        
        If userSelection = 2 Then                    'Canceled
            abort = True
            Exit Function
        End If
        
    End If
    
    On Error Resume Next
    result = (ThisWorkbook.VBProject.VBE.VBProjects.Count) > 0
    On Error GoTo 0
    
    CheckIfTrusted = result
    
End Function

'@Ignore ProcedureNotUsed
Public Sub CheckVBATrustedState()
    If Not IsTrustVbaSet Then
        Debug.Print "VBA 'IS NOT Trusted'"
    Else
        Debug.Print "VBA 'IS Trusted'"
    End If
End Sub


'@Ignore ProcedureNotUsed
Public Sub PrintAllReferences(ByVal project As Object) ' "ReferenceResolver.PrintAllReferences thisworkbook.VBProject" <- run to get values in Debug window
    Dim ref As Object
    For Each ref In project.References
        If Not ref.BuiltIn Then
            PrintToImmediateWindow ref
        End If
    Next
End Sub

Private Sub PrintToImmediateWindow(ByVal reference As Object)
    Debug.Print "'" + DEBUG_REFERENCE_SEPARATOR
    Debug.Print "' GUID:                  " + reference.GUID
    Debug.Print "' Major Version Number:  " + CStr(reference.Major)
    Debug.Print "' Minor Version Number:  " + CStr(reference.Minor)
    Debug.Print "' FullPath:              " + reference.fullPath
    Debug.Print "' Name:                  " + reference.Name
    Debug.Print "' Description:           " + reference.Description
    Debug.Print "' BuiltIn:               " + CStr(reference.BuiltIn)
    Debug.Print "'" + DEBUG_REFERENCE_SEPARATOR
End Sub


