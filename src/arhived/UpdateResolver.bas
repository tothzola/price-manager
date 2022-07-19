Attribute VB_Name = "UpdateResolver"
'@Folder("System.Update")
Option Explicit
Option Private Module

Private RunAutoUpdateAt As Date
Private Const CMDTEXT As String = "<LIST>" & _
"<VIEWGUID>{6925B711-7BC0-4090-AB33-27D9659EDA5C}</VIEWGUID><LISTNAME>{8754B0B5-5710-4FFA-B002-DFD1099CE3D3}</LISTNAME><LISTWEB>" & _
"https://portal.freudenberg-pm.com/sites/sscbrasov/Masterdata/Masterdata/ZoltanToth/_vti_bin</LISTWEB><LISTSUBWEB></LISTSUBWEB>" & _
"<ROOTFOLDER>/sites/sscbrasov/Masterdata/Masterdata/ZoltanToth/CommonAddIns</ROOTFOLDER></LIST>"

'@EntryPoint
Public Sub ManualUpdate()
    If Not ThisWorkbook.ReadOnly Then
        Application.StatusBar = "Checking (manualy) for updates from 'https://...'"
        CheckUpdate True
        Application.StatusBar = vbNullString
    End If
End Sub

'@EntryPoint
Public Sub AutoUpdate(Optional ByRef SkipProcedureForUpdate As Boolean)
    If Not ThisWorkbook.ReadOnly And Not Application.ProtectedViewWindows.Count > 0 Then
        Application.StatusBar = "checking (automat.) for updates from 'https://...'"
        CheckUpdate False, SkipProcedureForUpdate
        Application.StatusBar = vbNullString
    End If
    
End Sub

Private Sub CheckUpdate(ByVal manual As Boolean, Optional ByRef SkipProcedureForUpdate As Boolean)
    
    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    
    Dim Updater As IUpdateAddIn
    Set Updater = AddInManager.Create(CMDTEXT, manual)

    Dim OlderThen As Long: OlderThen = VBA.Now - Updater.GetLastUpdate

    If (OlderThen >= 7) Or manual Then
        Updater.IsThereAnUpdate SkipProcedureForUpdate
    End If

CleanExit:
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Set Updater = Nothing
    Exit Sub

CleanFail:
    MsgBox Err.Number & vbTab & Err.Description, vbCritical, Title:=SIGN
    LogManager.Log ErrorLevel, "Error: " & VBA.Err.Number & ". " & VBA.Err.Description & ". " & "Public Sub CheckAndUpdate()"
    Resume CleanExit
    Resume
    
End Sub


