VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Mathieu Guindon CR https://codereview.stackexchange.com/q/52185/146810"

'@Folder("Logger.TextWriter")
'@ModuleDescription("Mathieu Guindon CR https://codereview.stackexchange.com/q/52185/146810")
'@PredeclaredID
'@Exposed
Option Explicit

'manage opened files in a Scripting.Dictionary :
Private openedFiles As Scripting.Dictionary

'skip Scripting.Dictionary  lookup if only 1 file is opened:
Private quickWriteFile As Long

'@Ignore ProcedureNotUsed
Public Property Get OpenedFilesCount() As Long
    OpenedFilesCount = openedFiles.Count
End Property

Public Function OpenFile(ByVal FileName As String, Optional ByVal overwrite As Boolean = False) As Boolean

    Dim fileNumber As Long
    fileNumber = GetFileNumber(FileName)

    'guard against opening a file that's already opened:
    If fileNumber <> FreeFile Then
        OpenFile = True
        Exit Function
    End If

    On Error GoTo Catch

    If overwrite Or Dir(FileName) = VBA.Constants.vbNullString Then
        Open FileName For Output As #fileNumber
    Else
        Open FileName For Append As #fileNumber
    End If

    openedFiles.Add FileName, fileNumber
    quickWriteFile = IIf(openedFiles.Count = 1, fileNumber, 0)

Catch:
    If Err.Number <> 0 Then
        Err.Clear
    End If

    OpenFile = (openedFiles.exists(FileName))

End Function

Public Sub WriteLine(ByVal Data As String, Optional ByVal FileName As String = VBA.Constants.vbNullString)
    
    Dim fileNumber As Long

    On Error GoTo Catch

    If CanWrite(FileName, fileNumber) Then
        Print #fileNumber, Data
    Else
        Err.Raise InvalidFileName, VBA.Information.TypeName(Me) & ".WriteLine", "Invalid filename."
    End If

Catch:
    If Err.Number <> 0 Then
        RemoveByValue fileNumber, openedFiles
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If

End Sub

Private Function CanWrite(ByVal FileName As String, ByRef outFileNumber As Long) As Boolean
    Dim result As Boolean: result = False
    Dim fileNumber As Long
    
    If quickWriteFile <> 0 And FileName = VBA.Constants.vbNullString Then
        fileNumber = quickWriteFile
        CanWrite = True
    Else
        CanWrite = TryGetValue(FileName, result, openedFiles)
    End If

    outFileNumber = fileNumber

End Function

Public Sub CloseFile(Optional ByVal FileName As String = VBA.Constants.vbNullString)
    If openedFiles.Count = 0 Then Exit Sub

    Dim fileNumber As Long
    fileNumber = GetFileNumber(FileName)

    If fileNumber <> FreeFile Then
        Close #fileNumber
        RemoveByValue fileNumber, openedFiles
        If fileNumber = quickWriteFile Then quickWriteFile = 0
    End If
    
    If openedFiles.Count = 1 Then quickWriteFile = openedFiles.Items(0)

End Sub

Public Sub CloseAllFiles()
    Dim File As Variant
    For Each File In openedFiles.Items
        Close #File
    Next
    openedFiles.RemoveAll
    quickWriteFile = 0
End Sub

Private Function GetFileNumber(ByVal FileName As String) As Long
    Dim result As Long

    If quickWriteFile <> 0 And FileName = VBA.Constants.vbNullString Then
        result = quickWriteFile
    ElseIf Not TryGetValue(FileName, result, openedFiles) Then
        result = FreeFile
    End If

    GetFileNumber = result
End Function

Private Function TryGetValue(ByVal Key As Variant, ByRef outValue As Variant, ByVal dict As Scripting.Dictionary) As Boolean

    If dict.exists(Key) Then
        LetSet outValue, dict.Item(Key)
        TryGetValue = True
    Else
        TryGetValue = False
    End If

End Function

Private Sub RemoveByValue(ByVal lookupVal As Variant, ByVal dict As Scripting.Dictionary)

    Dim Key As Variant
    For Each Key In dict.Keys
        If dict.Item(Key) = lookupVal Then
            dict.Remove Key
            Exit For
        End If
    Next
    
End Sub

'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub LetSet(ByRef variable As Variant, ByVal Value As Variant)

    If IsObject(Value) Then
        Set variable = Value
    Else
        variable = Value
    End If
    
End Sub

Public Function Create() As TextWriter
    Set Create = New TextWriter
End Function

Private Sub Class_Initialize()
    Set openedFiles = New Scripting.Dictionary
End Sub

Private Sub Class_Terminate()
    CloseAllFiles
    Set openedFiles = Nothing
End Sub

