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
Implements IDisposable

Private Const salted As String = "D7D05"

Private Type TSHA256Managed
    Base As HashAlgorithmBase
    Disposed As Boolean
End Type

Private this As TSHA256Managed

'@Description("Factory method creates and returns a new instance of this class.")
Public Function Create() As IHashAlgorithm
Attribute Create.VB_Description = "Factory method creates and returns a new instance of this class."
    Set Create = New SHA256Managed
End Function

Private Sub Class_Initialize()
    Set this.Base = New HashAlgorithmBase
End Sub

Private Function IHashAlgorithm_ComputeHash(ByVal Value As String) As String
    Dim bytes() As Byte
    bytes = StrConv(salted & Value, vbFromUnicode)
    
    Dim Algo As Object
    Set Algo = CreateObject("System.Security.Cryptography.SHA256Managed")
    
    Dim buffer() As Byte
    buffer = Algo.ComputeHash_2(bytes)
        
    IHashAlgorithm_ComputeHash = this.Base.ToHexString(buffer)
    
End Function

Private Sub Dispose()

    If this.Disposed Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " instance was already disposed."
        Exit Sub
    End If

    If Not this.Base Is Nothing Then
        Disposable.TryDispose this.Base
        Set this.Base = Nothing
    End If
    
    this.Disposed = True
    
    #If TestMode Then
        LogManager.Log InfoLevel, VBA.Information.TypeName(Me) & " is terminating"
    #End If
    
End Sub

Private Sub Class_Terminate()
    If Not this.Disposed Then Dispose
End Sub

Private Sub IDisposable_Dispose()
    Dispose
End Sub
