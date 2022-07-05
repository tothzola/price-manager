VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AddInManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("System.Update.Impl")
'@PredeclaredId
Option Explicit

Implements IUpdateAddIn

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" ( _
    ByVal pCaller As LongPtr, _
    ByVal szURL As String, _
    ByVal szFileName As String, _
    ByVal dwReserved As LongPtr, _
    ByVal lpfnCB As LongPtr) As Long
    Private Declare PtrSafe Function DeleteUrlCacheEntry Lib "wininet.dll" _
    Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" _
                             Alias "URLDownloadToFileA" ( _
                             ByVal pCaller As Long, _
                             ByVal szURL As String, _
                             ByVal szFileName As String, _
                             ByVal dwReserved As Long, _
                             ByVal lpfnCB As Long) As Long
    Private Declare Function DeleteUrlCacheEntry Lib "wininet.dll" _
                             Alias "DeleteUrlCacheEntryA" (ByVal lpszUrlName As String) As Long
#End If

Private Const ERROR_SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10
Private Const FILE_LINK_TO_SHAREPOINT As String = "https://portal.freudenberg-pm.com/sites/sscbrasov/Masterdata/Masterdata/ZoltanToth/CommonAddIns/"

Private Type TAddInManager
    SPDOM As ISharePointDOM
    manual As Boolean
    version As String
    actualName As String
    tempName As String
    lastUpdate As Date
    hasPermission As Boolean
End Type

Private this As TAddInManager

Public Property Get AppName() As String
    AppName = APP_ADDIN_NAME
End Property

Public Property Get AppDownloadUrl() As String
    AppDownloadUrl = FILE_LINK_TO_SHAREPOINT & AppName
End Property

Public Property Get SPDOM() As ISharePointDOM
    Set SPDOM = this.SPDOM
End Property

Public Property Set SPDOM(ByVal RHS As ISharePointDOM)
    Set this.SPDOM = RHS
End Property

Public Property Get manual() As Boolean
    manual = this.manual
End Property

Public Property Let manual(ByVal RHS As Boolean)
    this.manual = RHS
End Property

Public Property Get version() As String
    version = this.version
End Property

Public Property Let version(ByVal RHS As String)
    this.version = RHS
End Property

Public Property Get actualName() As String
    actualName = this.actualName
End Property

Public Property Let actualName(ByVal RHS As String)
    this.actualName = RHS
End Property

Public Property Get tempName() As String
    tempName = this.tempName
End Property

Public Property Let tempName(ByVal RHS As String)
    this.tempName = RHS
End Property

Public Property Get lastUpdate() As Date

    On Error Resume Next
    this.lastUpdate = CDate(VBA.GetSetting(AppName, "Updates", "LastUpdate", "0"))
    If Not VBA.Err = 0 Then this.lastUpdate = 0
    
    If this.lastUpdate = 0 Then
        Dim dateNow As Date: dateNow = Int(VBA.Now)
        'Never checked for an update, save today!
        VBA.SaveSetting AppName, "Updates", "LastUpdate", CStr(Int(dateNow))
    End If
    lastUpdate = this.lastUpdate
    
End Property

Public Property Let lastUpdate(ByVal lastDate As Date)
    
    Dim tmpDate As Date
    tmpDate = VBA.Format$(VBA.Date, GetDateFormat)
    
    lastDate = tmpDate
    this.lastUpdate = lastDate
    
    VBA.SaveSetting AppName, "Updates", "LastUpdate", CStr(Int(this.lastUpdate))
    
End Property

'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As AddInManager
    Set Class = AddInManager
End Property

Public Function Create(ByVal Connection As String, Optional ByVal manual As Boolean) As IUpdateAddIn
    Guard.NonDefaultInstance Me
    Guard.EmptyString Connection
    
    Dim result As AddInManager
    Set result = New AddInManager

    Set result.SPDOM = SharePointDOM.Create(Connection)
    result.manual = manual
    result.version = ThisWorkbook.Names("MVVM_PRICEAPPROVAL_APPVERSION").RefersToRange.Value
    result.actualName = ThisWorkbook.FullName
        
    Set Create = result
    
End Function

Private Function DownloadFile(ByVal WebFilename As String, ByVal SaveFileName As String) As Boolean

    'Attempt to delete any cached version of the file. Since we're only interested in
    'nuking the file, the routine is called as a sub. If the return value is requires
    '(calling as a function), DeleteUrlCacheEntry returns 1 if successful, or 0 otherwise, e.g.
    'Note that the remote URL is passed as this is the name the cached file is known by. This does NOT delete the file from the remote server.

    If DeleteUrlCacheEntry(WebFilename) = 1 Then
        Debug.Print "cached file found and deleted"
    Else
        Debug.Print "no cached file for " & WebFilename
    End If

    'Download the file. BINDF_GETNEWESTVERSION forces the API to download from the specified source.
    'Passing 0& as dwReserved causes the locally-cached copy to be downloaded, if available.
    'If the API returns ERROR_SUCCESS (0), DownloadFile returns True.
    DownloadFile = URLDownloadToFile(0&, WebFilename, SaveFileName, BINDF_GETNEWESTVERSION, 0&) = ERROR_SUCCESS
    
End Function

Private Function CheckForNewVersion(ByRef outVersion As String) As Boolean

    Dim IsConnected As Boolean
    IsConnected = this.SPDOM.IsConnected
    
    Dim isNewVersion As Boolean
    isNewVersion = this.SPDOM.GetAddInVersion(APP_ADDIN_NAME, outVersion)
    
    If Not outVersion = vbNullString Then
        Dim majorVer As Variant
        majorVer = Split(outVersion, ".")
        outVersion = majorVer(0)
        
    End If
    
    this.hasPermission = isNewVersion
    
    CheckForNewVersion = IsConnected And isNewVersion

End Function

Private Sub RemoveOldCopy()

    this.tempName = this.actualName & "(OldVersion)"
    On Error Resume Next
    Kill this.tempName
    On Error GoTo 0
    
End Sub

Private Function GetUpdate() As Boolean
    
    On Error Resume Next
    'If workbook has been saved readonly, we can safely delete the file!
    If ThisWorkbook.ReadOnly Then
        Err.Clear
        Kill this.actualName
    End If
    lastUpdate = Now
    ThisWorkbook.SaveAs this.tempName
    DoEvents
    Kill this.actualName
    On Error GoTo 0
    If DownloadFile(AppDownloadUrl, this.actualName) And VBA.Err = 0 Then GetUpdate = True
    
End Function

Private Sub IUpdateAddIn_IsThereAnUpdate(Optional ByRef GetAppUpdates As Boolean)

    Dim newVersion As String
    If CheckForNewVersion(newVersion) Then
        
        If CLng(this.version) <> CLng(newVersion) Then
        
            Dim userChoiseIs As Long
            userChoiseIs = MsgBox("We have an update, do you wish to download?", vbQuestion + vbYesNo)
            
            If userChoiseIs = VBA.VbMsgBoxResult.vbYes Then
                    
                RemoveOldCopy
                            
                If GetUpdate Then
                    MsgBox "Successfully updated the addin," & vbCr _
                         & "please restart Excel to start using the new version!", vbOKOnly + vbInformation, "Success"
                    GetAppUpdates = True
                    Exit Sub
                Else
                    MsgBox "Updating has failed.", vbExclamation + vbOKOnly, "Failed"
                    Exit Sub
                End If
                
            End If
        End If
        
    ElseIf this.hasPermission And this.manual Then
        MsgBox "Your program is up to date.", vbInformation + vbOKOnly, "Up to date"
    End If
    
End Sub

Private Function IUpdateAddIn_GetLastUpdate() As Date
    IUpdateAddIn_GetLastUpdate = lastUpdate
    
End Function


