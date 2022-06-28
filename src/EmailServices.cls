VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EmailServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("AppObjects.Services")
Option Explicit

Private Const DEVELOPERS As String = "Zoltan.To@outlook.com;kamal.9328093207@gmail.com"

Private Enum olImportance
    olI_Low = 0
    olI_Normal = 1
    olI_High = 2
End Enum

Private Enum olSensivity
    olS_Normal = 0
    olS_Personal = 1
    olS_Private = 2
    olS_Confidential = 3
End Enum

'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As EmailServices
Attribute Class.VB_Description = "Returns class reference"
    Set Class = EmailServices
End Property

Private Sub Class_Initialize()
    Guard.Singleton Me
End Sub

Public Sub EmailNotifyUser(ByVal mailTo As String, _
                        ByVal mailSubject As String, _
                        ByVal mailDisplay As Boolean, _
                        ParamArray mailBody() As Variant)

    Dim body As String
    body = "<html><body style=""font-size:14px;font-family:Arial"">Hello, <br><br>"
    
    Dim i As Long
    For i = LBound(mailBody) To UBound(mailBody)
        If i = 0 Then
            body = body & mailBody(i) & "<br>"
            body = body & "<style> table.schedule td {border-collapse:collapse;border:1px solid black;}" _
                        & "<table class=""schedule"" style=""width:250px;""><tbody style=""font-size:12px;"">" _
                        & "<table class=""schedule"" style=""width:250px;""><tbody style=""font-size:12px;"">"
        
        ElseIf i = UBound(mailBody) Then
            body = body & "</td></tr></tbody></table><br>" & mailBody(i) & "<br><br></body></html>"
            
        Else
            body = body & "<tr><td>" & mailBody(i) & "</td></tr>"
            
        End If
    Next i
    
    AssambleEmail mailTo:=mailTo, _
                  mailSubject:=mailSubject, _
                  mailDisplay:=mailDisplay, _
                  mailBody:=body

End Sub

Public Sub EmailFeedback()

    Dim sendTo As String: sendTo = DEVELOPERS
    Dim subject As String: subject = "Price Approval App - Feedback"
    Dim body As String: body = "<BODY style=font-size:10.5pt;font-family:Arial>Hello Dev Team,</BODY>" & _
        "<BODY style=font-size:10.5pt;font-family:Arial><p> we would like you feedback to improve this app. </BODY>" & _
        "<BODY style=font-size:10.5pt;font-family:Arial><p> What is your opinion of this app? </BODY>" & _
        "<BODY style=font-size:10.5pt;font-family:Arial><p> Please leave your feedback below: </BODY>" & "<br><br>"
        
    AssambleEmail mailTo:=sendTo, _
                  mailSubject:=subject, _
                  mailBody:=body
End Sub

Public Sub EmailRecordRejected(ByVal UserName As String, _
                       ByVal mailTo As String, _
                       ByVal mailSubject As String, _
                       ByVal mailDisplay As Boolean, _
                       ParamArray mailBody() As Variant)
                            
    Dim body As String
    body = "<html><body style=""font-size:14px;font-family:Arial"">Hello " & UserName & ",<br><br>"

    Dim i As Long
    For i = LBound(mailBody) To UBound(mailBody)
        If i = 0 Then
            body = body & mailBody(i) & "<br>"
            body = body & "<style> table.schedule td {border-collapse:collapse;border:1px solid black;}" _
                        & "<table class=""schedule"" style=""width:250px;""><tbody style=""font-size:12px;"">" _
                        & "<table class=""schedule"" style=""width:250px;""><tbody style=""font-size:12px;"">"
            
        Else
        
            body = body & "<tr><td>" & mailBody(i) & "</td></tr>"
            
        End If
    Next i
    body = body & "</tbody></table><br><br></body></html>"
    
    AssambleEmail mailTo:=mailTo, _
                  mailSubject:=mailSubject, _
                  mailDisplay:=mailDisplay, _
                  mailBody:=body
                        
End Sub

Private Sub AssambleEmail(ByVal mailTo As String, ByVal mailSubject As String, _
                          Optional ByVal mailToCC As String = vbNullString, _
                          Optional ByVal mailBody As String = vbNullString, _
                          Optional ByVal mailDisplay As Boolean = True, _
                          Optional ByVal setImportance As olImportance = olI_Normal, _
                          Optional ByVal setSensivity As olSensivity = olS_Normal)
                          
    Dim OutlookApp As Outlook.Application
    Set OutlookApp = New Outlook.Application
    
    Dim ItemMail As Outlook.MailItem
    Set ItemMail = OutlookApp.CreateItem(Outlook.OlItemType.olMailItem)
    
    With ItemMail
        .To = mailTo
        .CC = mailToCC
        .subject = mailSubject
        
        Dim signature As String
        If GetSignature(signature) Then
            .HTMLBody = mailBody & signature
        Else
            .HTMLBody = mailBody
        End If
        
        .Importance = setImportance
        .Sensitivity = setSensivity
        If mailDisplay Then
            .Display
        Else
            .Send
        End If
    End With
    
    Set OutlookApp = Nothing
    Set ItemMail = Nothing
        
End Sub

Private Function GetSignature(ByRef outSign As String) As Boolean

    On Error GoTo CleanFail
    
    Dim EnvPath As String
    EnvPath = VBA.Environ$("APPDATA") & "\Microsoft\Signatures\"

    Dim signPath As String                       'get First Signature file
    signPath = VBA.Dir(EnvPath & "*.htm")
    
    If Not signPath = vbNullString Then
        Dim imagePath As String                  'get First Signature file pictures Folder
        imagePath = VBA.Dir(EnvPath & VBA.Mid$(signPath, 1, VBA.InStr(1, signPath, ".") - 1) & "*", vbDirectory)
                            
        If imagePath = signPath Then
            imagePath = VBA.Dir(EnvPath & VBA.Mid$(signPath, 1, VBA.InStr(1, signPath, ".") - 1) & "_*", vbDirectory)
        End If
    End If
    
    Dim signFolderPath As String: signFolderPath = EnvPath & imagePath
    
    Dim TempString As String
    If VBA.Dir(EnvPath & signPath) <> vbNullString And GetTextStream(EnvPath & signPath, TempString) Then
    
        outSign = VBA.Replace(TempString, imagePath, signFolderPath)
        Dim result As Boolean: result = Not (outSign = vbNullString)
        
    End If
    
CleanExit:
    GetSignature = result
    Exit Function

CleanFail:
    MsgBox VBA.Err.Description, vbCritical, Title:=SIGN
    LogManager.Log ErrorLevel, VBA.TypeName(Me) & " Error: " & VBA.Err.Number & ". " & VBA.Err.Description
    Resume CleanExit
    Resume
    
End Function

Private Function GetTextStream(ByVal strPath As String, ByRef outString As String) As Boolean

    Dim SystemFile As Scripting.FileSystemObject
    Set SystemFile = New Scripting.FileSystemObject
    
    Dim Stream As Scripting.TextStream
    On Error Resume Next
    Set Stream = SystemFile.GetFile(strPath).OpenAsTextStream(1, -2)
    On Error GoTo 0
    
    If Not Stream Is Nothing Then
        outString = Stream.ReadAll
        Stream.Close
        GetTextStream = outString <> vbNullString
    End If
    
    Set Stream = Nothing
    Set SystemFile = Nothing
    
End Function

