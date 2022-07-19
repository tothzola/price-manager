VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SharePointDOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "Rubberduck" ,"Predeclared Class Module"
'@Folder("System.Update.Impl")
'@PredeclaredId
Option Explicit

Implements ISharePointDOM

#If VBA7 And Win64 Then
    Private Stat As LongPtr
    Private Declare PtrSafe Function InternetGetConnectedState _
    Lib "wininet.dll" (lpdwFlags As LongPtr, _
    ByVal dwReserved As Long) As Boolean
#Else
    Private Stat As Long
    Private Declare Function InternetGetConnectedState _
                             Lib "wininet.dll" (lpdwFlags As Long, _
                                                ByVal dwReserved As Long) As Boolean
#End If

Private Type TState
    Connection As String
    AddIns As Scripting.Dictionary
    DOMDocument As MSXML2.DOMDocument60
    XMLHTTP As MSXML2.XMLHTTP60
    ListWeb As String
    ListGUID As String
    ViewGUID As String

End Type

Private this As TState

Public Property Get Connection() As String
    Connection = this.Connection
End Property

Public Property Let Connection(ByVal RHS As String)
    this.Connection = RHS
End Property

'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As SharePointDOM
    Set Class = SharePointDOM
End Property

Public Function Create(ByVal Connection As String) As ISharePointDOM
    Guard.NonDefaultInstance Me
    Guard.EmptyString Connection
    
    Dim result As SharePointDOM
    Set result = New SharePointDOM
    
    result.Connection = Connection
    
    Set Create = result
    
End Function

Private Sub Class_Initialize()
    Set this.AddIns = New Scripting.Dictionary
    Set this.DOMDocument = New MSXML2.DOMDocument60
    Set this.XMLHTTP = New MSXML2.XMLHTTP60
    
End Sub

Private Function ResolveVersion(ByVal Name As String, ByRef outVersion As String) As Boolean
    
    Dim result As Boolean
    On Error GoTo CleanFail
    
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
    
    Set this.DOMDocument = this.XMLHTTP.responseXML
    this.DOMDocument.SetProperty "SelectionLanguage", "XPath"
    
    Dim viewNodesFilename As IXMLDOMNodeList
    Set viewNodesFilename = this.DOMDocument.DocumentElement.SelectNodes("//*/*/@ows_LinkFilename")
    
    Dim viewNodesUIVersion As IXMLDOMNodeList
    Set viewNodesUIVersion = this.DOMDocument.DocumentElement.SelectNodes("//*/*/@ows__UIVersionString")
        
    Dim i As Long
    For i = 1 To viewNodesFilename.Length
        this.AddIns.Add Key:=viewNodesFilename.Item(i - 1).NodeValue, Item:=viewNodesUIVersion.Item(i - 1).NodeValue
    Next
    
    If this.AddIns.exists(Name) Then
        outVersion = this.AddIns.Item(Name)
        result = Not outVersion
    End If
    
CleanExit:
    ResolveVersion = result
    Exit Function
    
CleanFail:
    If Err.Description Like "Access is denied*" Then
        'MsgBox "Could not acces version information!", vbExclamation + vbOKOnly, SIGN
    End If
    Resume CleanExit
    Resume
    
End Function

Private Function ISharePointDOM_IsConnected() As Boolean
    ISharePointDOM_IsConnected = (InternetGetConnectedState(Stat, 0&) <> 0)
End Function

Private Function ISharePointDOM_GetAddInVersion(ByVal Name As String, ByRef outVersion As String) As Boolean
    ISharePointDOM_GetAddInVersion = ResolveVersion(Name, outVersion)
End Function


