VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SHA256Managed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Authentication.Hashing")
'@PredeclaredId
Option Explicit
Implements IHashAlgorithm

Private Const salted As String = "D7D05"

Private Base As HashAlgorithmBase

'@Description("Factory method creates and returns a new instance of this class.")
Public Function Create() As IHashAlgorithm
    Set Create = New SHA256Managed
End Function

Private Sub Class_Initialize()
    Set Base = New HashAlgorithmBase
End Sub

Private Function IHashAlgorithm_ComputeHash(ByVal Value As String) As String
    Dim bytes() As Byte
    bytes = StrConv(salted & Value, vbFromUnicode)
    
    Dim Algo As Object
    Set Algo = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    Dim buffer() As Byte
    buffer = Algo.ComputeHash_2(bytes)
        
    IHashAlgorithm_ComputeHash = Base.ToHexString(buffer)
End Function
