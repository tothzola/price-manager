Attribute VB_Name = "Templates"
'@Folder("Common")
Option Explicit
Option Private Module

Private Const DEVELOPERS As String = "Zoltan.To@outlook.com;kamal.9328093207@gmail.com"

Public Enum olImportance
    olI_Low = 0
    olI_Normal = 1
    olI_High = 2
End Enum

Public Enum olSensivity
    olS_Normal = 0
    olS_Personal = 1
    olS_Private = 2
    olS_Confidential = 3
End Enum

Private Sub AssambleEmail(ByVal mailTo As String, ByVal mailSubject As String, _
                          Optional ByVal mailToCC As String = vbNullString, _
                          Optional ByVal mailBody As String = vbNullString, _
                          Optional ByVal setImportance As olImportance = olI_Normal, _
                          Optional ByVal setSensivity As olSensivity = olS_Normal)
    
    Dim OutlookApp As Outlook.Application
    Set OutlookApp = New Outlook.Application
    
    Dim CreateMail As Outlook.MailItem
    Set CreateMail = OutlookApp.CreateItem(Outlook.OlItemType.olMailItem)

    If Not CreateMail Is Nothing Then

        With CreateMail
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
            .Display
        End With
    
    End If
    
    Set OutlookApp = Nothing
    Set CreateMail = Nothing
    
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

Public Sub EmailFeedback()

    Dim sendTo As String: sendTo = DEVELOPERS
    Dim subject As String: subject = "Price Approval Manager - Feedback"
    Dim body As String: body = "<BODY style=font-size:10.5pt;font-family:Arial>Hello Dev Team,</BODY>" & _
        "<BODY style=font-size:10.5pt;font-family:Arial><p> we would like you feedback to improve this app. </BODY>" & _
        "<BODY style=font-size:10.5pt;font-family:Arial><p> What is your opinion of this app? </BODY>" & _
        "<BODY style=font-size:10.5pt;font-family:Arial><p> Please leave your feedback below: </BODY>" & "<br><br>"
        
    AssambleEmail mailTo:=sendTo, _
                    mailSubject:=subject, _
                    mailBody:=body

End Sub

