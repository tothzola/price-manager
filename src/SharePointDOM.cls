VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SharePointDOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("System.Updater")

Option Explicit

Private Type TState
    Connection As String
    FilesFound As Scripting.Dictionary
    DOMDocument As MSXML2.DOMDocument60
    XMLHTTP As MSXML2.XMLHTTP60
    ListWeb As String
    ListGUID As String
    ViewGUID As String

End Type

Private this As TState

Private Sub Class_Initialize()
    Set this.FilesFound = New Scripting.Dictionary
    Set this.DOMDocument = New MSXML2.DOMDocument60
    Set this.XMLHTTP = New MSXML2.XMLHTTP60
End Sub

Public Function HasAcces(ByVal Connection As String) As Boolean
    
    Dim result As Boolean
    On Error GoTo CleanFail
    
    If Not System.ConnectedToNetwork Then GoTo CleanFail
    
    this.DOMDocument.LoadXML Connection
    this.ViewGUID = this.DOMDocument.SelectSingleNode("//*/VIEWGUID").Text
    this.ListGUID = this.DOMDocument.SelectSingleNode("//*/LISTNAME").Text
    this.ListWeb = this.DOMDocument.SelectSingleNode("//*/LISTWEB").Text

    Dim URL As String                            'add dummy param dt and set it to the current date/time so the result will not be cached.
    URL = this.ListWeb & "/owssvr.dll?Cmd=Display&List=" _
        & this.ListGUID & "&View=" _
        & this.ViewGUID & "&XMLDATA=TRUE&dt=" _
        & Now
    
    this.XMLHTTP.Open "GET", URL, varAsync:=False, bstrUser:=vbNullString, bstrPassword:=vbNullString
    this.XMLHTTP.send                            'make the call and get the response from the server
    
    result = (this.XMLHTTP.Status = 200)
    
CleanExit:
    HasAcces = result
    Exit Function
    
CleanFail:
    MsgBox "Could not acces SharePoint!", vbExclamation + vbOKOnly, PriceApprovalSignature
    Resume CleanExit
    Resume
    
End Function

Public Function GetVersion(ByVal Connection As String, ByVal Name As String) As String
    
    Dim result As String
    On Error GoTo CleanFail
    
    If Not HasAcces(Connection) Then GoTo CleanFail

    Set this.DOMDocument = this.XMLHTTP.responseXML
    this.DOMDocument.SetProperty "SelectionLanguage", "XPath"
    
    Dim nodesFilename As IXMLDOMNodeList
    Set nodesFilename = this.DOMDocument.DocumentElement.SelectNodes("//*/*/@ows_LinkFilename")
    
    Dim nodesUIVersion As IXMLDOMNodeList
    Set nodesUIVersion = this.DOMDocument.DocumentElement.SelectNodes("//*/*/@ows__UIVersionString")
        
    Dim i As Long
    For i = 1 To nodesFilename.Length
        this.FilesFound.Add Key:=nodesFilename.Item(i - 1).NodeValue, Item:=nodesUIVersion.Item(i - 1).NodeValue
    Next
    
    If this.FilesFound.exists(Name) Then
        result = this.FilesFound.Item(Name)
    End If
    
CleanExit:
    GetVersion = result
    this.FilesFound.RemoveAll

    Exit Function
    
CleanFail:
    If Err.Description Like "Access is denied*" Then
        MsgBox "Could not acces version information!", vbExclamation + vbOKOnly, PriceApprovalSignature
    End If
    Resume CleanExit
    Resume
    
End Function
