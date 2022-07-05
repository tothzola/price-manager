VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDisposable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents an object that can be (or must be) explicitly terminated."
'@ModuleDescription "Represents an object that can be (or must be) explicitly terminated."
'@Folder("MVP.Model")
'@Interface
'@Exposed
'Copyright (c) 2020 Mathieu Guindon
Option Explicit

'@Description "Releases references held, prepares the object for eventual destruction."
Public Sub Dispose()
Attribute Dispose.VB_Description = "Releases references held, prepares the object for eventual destruction."
End Sub
