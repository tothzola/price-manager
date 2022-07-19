VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Default instance factory; avoid altering default instance state."
'@ModuleDescription "Default instance factory; avoid altering default instance state."
'@Folder("System.RibbonBackup")
'@PredeclaredId
'@Exposed
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As LongPtr)
#Else
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
#End If

Private Type TObjectBackup

    #If VBA7 Then
    Pointer As LongPtr
    #Else
    Pointer As Long
    #End If
    
    NameManager As Excel.Names
    
End Type

Private Const APP_RIBBONPOINTER As String = "PRICEAPPROVAL_RIBBONPOINTER"
Private this As TObjectBackup

'@Description("Adds Object instance to Excel Name Manager for later use")
Public Sub AddObject(ByVal Instance As Object)
Attribute AddObject.VB_Description = "Adds Object instance to Excel Name Manager for later use"
    With New ObjectBackup
        'APP_RIBBONPOINTER name must be added to the Name Manager before this object is instanciated.
        .NameManager.Add Name:=APP_RIBBONPOINTER, RefersTo:=VBA.ObjPtr(Instance)
        ThisWorkbook.Saved = True
    End With
End Sub


Friend Property Get NameManager() As Excel.Names
    Dim Book As Excel.Workbook
    Set Book = ThisWorkbook
    Set NameManager = Book.Names
End Property


'@Description("Gets Object instance from Excel Name Manager")
Public Property Get GetObject() As Object
Attribute GetObject.VB_Description = "Gets Object instance from Excel Name Manager"

    On Error GoTo CleanFail
    
    #If VBA7 Then
        this.Pointer = CLngPtr(VBA.Strings.Replace(NameManager.Item(APP_RIBBONPOINTER).Value, "=", VBA.Constants.vbNullString))
    #Else
        this.Pointer = CLng(VBA.Strings.Replace(NameManager.Item(APP_RIBBONPOINTER).Value, "=", VBA.Constants.vbNullString))
    #End If
    
    #If VBA7 Then
        Dim pointerSize As LongPtr
    #Else
        Dim pointerSize As Long
    #End If
    
    pointerSize = VBA.Strings.LenB(this.Pointer)
    
    'On Error System crash, excel workbook closes, cannot recover, use this function whisely.
    Dim obj As Object
    CopyMemory obj, this.Pointer, pointerSize
    Set GetObject = obj
    
CleanExit:
    CopyMemory obj, 0&, pointerSize
    Exit Property
    
CleanFail:
    Set GetObject = Nothing
    Resume CleanExit
    Resume
    
End Property
