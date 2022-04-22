VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWeakReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes an object that holds the address of a pointer to another object."
'@Folder("WeakReference")
'@ModuleDescription("Describes an object that holds the address of a pointer to another object.")
'@Interface
'@Exposed
Option Explicit

'@Description "Gets the ribbon object at the held pointer. Pointer is Named Range in ThisWorkbook"
Public Property Get Ribbon() As Object
Attribute Ribbon.VB_Description = "Gets the ribbon object at the held pointer. Pointer is Named Range in ThisWorkbook"
End Property
