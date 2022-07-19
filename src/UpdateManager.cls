VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UpdateManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("System.Updater")
'@PredeclaredId
Option Explicit

Implements IUpdate

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

Private Enum u__Notificate
    u__Update
    u__UpToDate
    u__Successful
    u__Intalled
    u__Skipped
    u__Failed
End Enum

Private Const AUTO_UPDATE_LIMIT As Long = 7 'Days
Private Const FILE_LINK_TO_SHAREPOINT As String = "https://SharepointSite"
Private Const SUCCESS As Long = 0
Private Const BINDF_GETNEWESTVERSION As Long = &H10

Private Type TAddInManager
    SharePoint As SharePointDOM
    Connection As String
    TempName As String
    LastUpdate As Date
    AppName As String
    ActualName As String
    NewVersion As Long
    UpToDate As Boolean
    Installed As Boolean
    Updated As Boolean
End Type

Private this As TAddInManager

Public Property Get AppName() As String
    AppName = this.AppName
End Property

Friend Property Let AppName(ByVal RHS As String)
    Guard.DefaultInstance Me
    this.AppName = RHS
End Property

Public Property Get AppDownloadUrl() As String
    AppDownloadUrl = FILE_LINK_TO_SHAREPOINT & AppName
End Property

Public Property Get ActualName() As String
    ActualName = this.ActualName
End Property

Friend Property Let ActualName(ByVal RHS As String)
    Guard.DefaultInstance Me
    this.ActualName = RHS
End Property

Public Property Get Connection() As String
    Connection = this.Connection
End Property

Friend Property Let Connection(ByVal RHS As String)
    Guard.DefaultInstance Me
    this.Connection = RHS
End Property

Friend Property Let NewVersion(ByVal RHS As Long)
    Guard.DefaultInstance Me
    this.NewVersion = RHS
End Property

Friend Property Let UpToDate(ByVal RHS As Boolean)
    Guard.DefaultInstance Me
    this.UpToDate = RHS
End Property

Friend Property Let Installed(ByVal RHS As Boolean)
    Guard.DefaultInstance Me
    this.Installed = RHS
End Property

Public Property Get LastUpdate() As Date

    On Error Resume Next
    'TEST: ?CDate(VBA.GetSetting("PriceApproval.xlam", "Updates", "LastUpdate", "0"))
    this.LastUpdate = CDate(VBA.GetSetting(AppName, "Updates", "LastUpdate", "0"))
    If Not VBA.Err = 0 Then this.LastUpdate = 0: VBA.Err.Clear
    
    If this.LastUpdate = 0 Then
        Dim dateNow As Date: dateNow = VBA.Int(VBA.Now)
        'Never checked for an update, save today!
        'TEST: VBA.SaveSetting "PriceApproval.xlam", "Updates", "LastUpdate", CStr(VBA.Int(VBA.Now-3))
        VBA.SaveSetting AppName, "Updates", "LastUpdate", CStr(Int(dateNow))
    End If
    LastUpdate = this.LastUpdate
    
End Property

Friend Property Let LastUpdate(ByVal lastDate As Date)
    Guard.DefaultInstance Me
    
    Dim tmpDate As Date
    tmpDate = VBA.Format$(VBA.Date, GetDateFormat)
    
    lastDate = tmpDate
    this.LastUpdate = lastDate
    
    VBA.SaveSetting AppName, "Updates", "LastUpdate", CStr(Int(this.LastUpdate))
    
End Property

Private Function NotifMessage(ByVal NotifType As u__Notificate) As String

    Select Case NotifType
        Case Is = u__Notificate.u__Update
            NotifMessage = "New Update available !" & vbNewLine & vbNewLine & "To update, click 'Ok' to continue."
            
        Case Is = u__Notificate.u__UpToDate
            NotifMessage = "This AddIn App is up to date." & vbNewLine _
                         & "Version: " & System.CurrentVersion
            
        Case Is = u__Notificate.u__Successful
            NotifMessage = "Updated Successful." & vbNewLine & vbNewLine _
                         & "Please Restart Excel Application to use the new AddIn version."
           
        Case Is = u__Notificate.u__Intalled
            NotifMessage = "Restart Excel Application to use the new AddIn version !"
            
        Case Is = u__Notificate.u__Skipped
            NotifMessage = "One critical update was skipped !" & vbNewLine & vbNewLine _
                         & "To use this App install the available Update."
        
        Case Is = u__Notificate.u__Failed
            NotifMessage = "An error occured while attempting to update this Application." & vbNewLine & vbNewLine _
                           & "Please contact " & DEVELOPER_NAME & " at " & DEVELOPER_EMAIL
        Case Else
            NotifMessage = vbNullString
            
    End Select
    
End Function

'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As UpdateManager
    Set Class = UpdateManager
End Property

Public Function Create(ByVal Connection As String, ByVal AppName As String) As IUpdate
    Guard.NonDefaultInstance UpdateManager
    Guard.EmptyString Connection
    
    Dim result As UpdateManager
    Set result = New UpdateManager
        result.Connection = Connection
        result.AppName = AppName
        result.ActualName = ThisWorkbook.FullName
        
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
    DownloadFile = URLDownloadToFile(0&, WebFilename, SaveFileName, BINDF_GETNEWESTVERSION, 0&) = SUCCESS
    
End Function

Private Function VersionFromSharepoint() As Long
    
    Dim result As Variant
    result = this.SharePoint.GetVersion(this.Connection, this.AppName)
    
    If Not result = vbNullString Then
        Dim getMajorVer As Variant
        getMajorVer = VBA.Split(result, ".")
        result = getMajorVer(0)
    Else
        result = System.CurrentVersion
    End If

    VersionFromSharepoint = CLng(result)

End Function

Private Function OldCopyRemoved() As Boolean

    this.TempName = this.ActualName & "(OldVersion)"
    On Error Resume Next
    Kill this.TempName
    OldCopyRemoved = VBA.Err = 0
    
End Function

Private Function DownloadSuccessful() As Boolean
    
    On Error Resume Next
    'If workbook has been saved readonly, we can safely delete the file!
    If ThisWorkbook.ReadOnly Then
        Err.Clear
        Kill this.ActualName
    End If
    
    LastUpdate = Now
    
    Application.DisplayAlerts = False
        DoSave = True
            ThisWorkbook.SaveAs this.TempName
        DoSave = False
    Application.DisplayAlerts = True
    
    DoEvents
    Kill this.ActualName

    DownloadSuccessful = DownloadFile(AppDownloadUrl, this.ActualName) And VBA.Err = 0
    
End Function

Private Sub Download(ByRef successful As Boolean)
    
    If Not this.SharePoint.HasAcces(this.Connection) Then Exit Sub
    
    If System.CurrentVersion <> this.NewVersion Then
    
        Dim oldRemoved As Boolean
        oldRemoved = OldCopyRemoved
        
        Dim newInstalled As Boolean
        newInstalled = DownloadSuccessful
        
        If newInstalled Then
            MsgBox NotifMessage(NotifType:=u__Successful), vbInformation + vbOKOnly, APP_SIGNATURE
            Installed = newInstalled
            successful = newInstalled
        Else
            MsgBox NotifMessage(NotifType:=u__Failed), vbCritical, APP_SIGNATURE
            LogManager.Log ErrorLevel, "Error: " & ErrNo.UpdateFailed & ". " & VBA.TypeName(Me) & ". " & "System Update Failed"
        End If
    Else
        MsgBox NotifMessage(NotifType:=u__UpToDate), vbInformation + vbOKOnly, APP_SIGNATURE
        UpToDate = Not successful
    End If

End Sub

Private Function UpdateAvailable() As Boolean

    NewVersion = VersionFromSharepoint
    
    Dim OlderThen As Long: OlderThen = VBA.Now - LastUpdate
    Dim timeToUpdate As Boolean: timeToUpdate = (OlderThen >= AUTO_UPDATE_LIMIT)
    
    If timeToUpdate Then
        UpToDate = (System.CurrentVersion = this.NewVersion)
        UpdateAvailable = (System.CurrentVersion <> this.NewVersion)
        
    Else
        UpToDate = Not timeToUpdate
        
    End If
    
End Function

Private Sub SystemUpdated(ByRef systemOk As Boolean)

    If this.Installed Then
        MsgBox NotifMessage(NotifType:=u__Intalled), vbExclamation + vbOKOnly, APP_SIGNATURE
        Exit Sub
        
    ElseIf Not this.UpToDate Then
        If MsgBox(NotifMessage(NotifType:=u__Update), vbExclamation + vbOKCancel, APP_SIGNATURE) = VBA.VbMsgBoxResult.vbOK Then
            Dim result As Boolean
            Download result
            UpToDate = result
            Exit Sub
        Else
            MsgBox NotifMessage(NotifType:=u__Skipped), vbOKOnly + vbCritical, APP_SIGNATURE
            Exit Sub
        End If
    Else
        result = this.UpToDate

    End If
    
    systemOk = result
    
End Sub

Private Sub Class_Initialize()
    Set this.SharePoint = New SharePointDOM
End Sub

Private Property Get IUpdate_NoUpdate() As Boolean
    Dim systemOk As Boolean
    SystemUpdated systemOk
    IUpdate_NoUpdate = systemOk
End Property

Private Sub IUpdate_Download()
    Dim successful As Boolean
    Download successful
End Sub

Private Function IUpdate_Available() As Boolean
    IUpdate_Available = UpdateAvailable()
End Function


