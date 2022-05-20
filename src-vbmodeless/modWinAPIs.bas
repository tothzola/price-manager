Attribute VB_Name = "modWinAPIs"
'@Folder "View"
Option Explicit
Option Private Module

'@Got this code from wellsr.com
'https://wellsr.com/vba/2016/excel/create-awesome-excel-splash-screen-for-your-spreadsheet/
'I then updated it for VBA7 and Else Version of VBA
Public Const GWL_STYLE = -16
Public Const WS_CAPTION = &HC00000

#If VBA7 Then
    Public Declare PtrSafe Function GetWindowLong _
                           Lib "user32" Alias "GetWindowLongA" ( _
                           ByVal hwnd As LongPtr, _
                           ByVal nIndex As LongPtr) As LongPtr
    Public Declare PtrSafe Function SetWindowLong _
                           Lib "user32" Alias "SetWindowLongA" ( _
                           ByVal hwnd As LongPtr, _
                           ByVal nIndex As LongPtr, _
                           ByVal dwNewLong As LongPtr) As LongPtr
    Public Declare PtrSafe Function DrawMenuBar _
                           Lib "user32" ( _
                           ByVal hwnd As LongPtr) As LongPtr
    Public Declare PtrSafe Function FindWindowA _
                           Lib "user32" (ByVal lpClassName As String, _
                           ByVal lpWindowName As String) As LongPtr
#Else
    Public Declare Function GetWindowLong _
                           Lib "user32" Alias "GetWindowLongA" ( _
                           ByVal hWnd As Long, _
                           ByVal nIndex As Long) As Long
    Public Declare Function SetWindowLong _
                           Lib "user32" Alias "SetWindowLongA" ( _
                           ByVal hWnd As Long, _
                           ByVal nIndex As Long, _
                           ByVal dwNewLong As Long) As Long
    Public Declare Function DrawMenuBar _
                           Lib "user32" ( _
                           ByVal hWnd As Long) As Long
    Public Declare Function FindWindowA _
                           Lib "user32" (ByVal lpClassName As String, _
                           ByVal lpWindowName As String) As Long
#End If

Sub HideTitleBar(frm As Object)
#If VBA7 Then
    Dim lngWindow As LongPtr
    Dim lFrmHdl As LongPtr
#Else
    Dim lngWindow As Long
    Dim lFrmHdl As Long
#End If
    lFrmHdl = FindWindowA(vbNullString, frm.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
End Sub

