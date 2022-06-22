VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WeakReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Default instance factory; avoid altering default instance state."
'@IgnoreModule ProcedureNotUsed, ImplicitUnboundDefaultMemberAccess, VariableNotAssigned, UnassignedVariableUsage
'@PredeclaredId
'@Exposed
'@Folder("WeakReference.Impl")
'@ModuleDescription("Default instance factory; avoid altering default instance state.")
Option Explicit

Implements IWeakReference

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As LongPtr)
#Else
    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
#End If

Private Type TReference

    #If VBA7 Then
    Pointer As LongPtr
    #Else
    Pointer As Long
    #End If
    
    WorkbookPointer As Names
    
End Type

Private this As TReference

'@Description("Default instance factory method.")
Public Function Create(ByVal Instance As Object) As IWeakReference
Attribute Create.VB_Description = "Default instance factory method."

    With New WeakReference
        .WorkbookPointer.Add Name:="PRICEAPPROVAL_RIBBONPOINTER", RefersTo:=VBA.ObjPtr(Instance)
        
        ThisWorkbook.Saved = True

        Set Create = .Self
    End With
    
End Function

Public Property Get Self() As IWeakReference
    Set Self = Me
End Property

Public Property Get WorkbookPointer() As Names
    Dim Book As Excel.Workbook
    Set Book = ThisWorkbook
    Set WorkbookPointer = Book.Names
End Property

Public Property Get Ribbon() As Object
    Set Ribbon = IWeakReference_Ribbon
End Property

Private Property Get IWeakReference_Ribbon() As Object

    On Error GoTo CleanFail
    
    #If VBA7 Then
        this.Pointer = CLngPtr(VBA.Strings.Replace(WorkbookPointer.Item("PRICEAPPROVAL_RIBBONPOINTER").Value, "=", VBA.Constants.vbNullString))
    #Else
        this.Pointer = CLng(VBA.Strings.Replace(WorkbookPointer.Item("PRICEAPPROVAL_RIBBONPOINTER").Value, "=", VBA.Constants.vbNullString))
    #End If
    
    #If VBA7 Then
        Dim pointerSize As LongPtr
    #Else
        Dim pointerSize As Long
    #End If
    
    pointerSize = VBA.Strings.LenB(this.Pointer)
    
    Dim obj As Object
    CopyMemory obj, this.Pointer, pointerSize
    
    Set IWeakReference_Ribbon = obj
    CopyMemory obj, 0&, pointerSize
    
CleanExit:
    Exit Property
    
CleanFail:
    MsgBox VBA.Err.Description, Title:=VBA.Err.Number
    LogManager.Log ErrorLevel, VBA.TypeName(Me) & " Error: " & VBA.Err.Number & ". " & VBA.Err.Description
    Set IWeakReference_Ribbon = Nothing
    Resume CleanExit
    Resume
    
End Property
