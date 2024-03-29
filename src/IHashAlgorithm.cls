VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IHashAlgorithm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "An interface representing a hashing Algorithm."
'@Folder("Authentication.Hashing")
'@ModuleDescription("An interface representing a hashing Algorithm.")
'@Interface
'@Exposed
Option Explicit

'@Description("Computes a hash for the given string value.")
Public Function ComputeHash(ByVal Value As String) As String
Attribute ComputeHash.VB_Description = "Computes a hash for the given string value."
End Function
