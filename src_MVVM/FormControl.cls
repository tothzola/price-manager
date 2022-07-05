VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("FormControl")
'@PredeclaredId
'@Exposed
Option Explicit
Option Compare Text

'By Chip Pearson, chip@cpearson.com, www.cpearson.com
'URL: http://www.cpearson.com/Excel/FormControl.aspx

Private Const C_USERFORM_CLASSNAME = "ThunderDFrame"
Private Const C_EXCEL_APP_CLASSNAME = "XLMain"
Private Const C_EXCEL_DESK_CLASSNAME = "XLDesk"
Private Const C_EXCEL_WINDOW_CLASSNAME = "Excel7"
Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000
Private Const MF_ENABLED = &H0&
Private Const MF_DISABLED = &H2&
Private Const MF_GRAYED = &H1&
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
Private Const GWL_HWNDPARENT = (-8)
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&
Private Const C_ALPHA_FULL_TRANSPARENT As Byte = 0
Private Const C_ALPHA_FULL_OPAQUE As Byte = 255
Private Const WS_DLGFRAME = &H400000
Private Const WS_THICKFRAME = &H40000
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000

Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const HKEY_CLASSES_ROOT  As Long = &H80000000
Private Const HKEY_CURRENT_CONFIG  As Long = &H80000005
Private Const HKEY_DYN_DATA  As Long = &H80000006
Private Const HKEY_PERFORMANCE_DATA  As Long = &H80000004
Private Const HKEY_USERS  As Long = &H80000003
Private Const KEY_ALL_ACCESS  As Long = &H3F
Private Const ERROR_SUCCESS  As Long = 0&
Private Const HKCU  As Long = HKEY_CURRENT_USER
Private Const HKLM  As Long = HKEY_LOCAL_MACHINE

Private Enum REG_DATA_TYPE
    REG_DATA_TYPE_DEFAULT = 0                    ' Default based on data type of value.
    REG_INVALID = -1                             ' Invalid
    REG_SZ = 1                                   ' String
    REG_DWORD = 4                                ' Long
End Enum

Public Enum FORM_PARENT_WINDOW_TYPE
    FORM_PARENT_NONE = 0
    FORM_PARENT_APPLICATION = 1
    FORM_PARENT_WINDOW = 2
End Enum

'''64 & 32-bit compatability
#If VBA7 Then
    Private Declare PtrSafe Function SetParent Lib "user32" ( _
    ByVal hWndChild As LongPtr, ByVal hWndNewParent As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hWnd As LongPtr, ByVal nIndex As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hWnd As LongPtr, ByVal nIndex As LongPtr, ByVal dwNewLong As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr

    Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
    
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function GetMenuItemCount Lib "user32" (ByVal hMenu As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function GetSystemMenu Lib "user32" (ByVal hWnd As LongPtr, ByVal bRevert As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function RemoveMenu Lib "user32" (ByVal hMenu As LongPtr, ByVal nPosition As LongPtr, ByVal wFlags As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As LongPtr) As LongPtr

    Private Declare PtrSafe Function EnableMenuItem Lib "user32" ( _
    ByVal hMenu As LongPtr, ByVal wIDEnableItem As LongPtr, ByVal wEnable As LongPtr) As LongPtr
    
    #If Win64 Then
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    #Else
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    #End If
    
    Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" ( _
    ByVal hWnd As LongPtr, ByVal lpString As String, ByVal cch As LongPtr) As Long
    
    Private Declare PtrSafe Function RegOpenKeyEx Lib "AdvApi32.dll" Alias "RegOpenKeyExA" ( _
    ByVal HKey As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As LongPtr, _
    ByVal samDesired As LongPtr, phkResult As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function RegQueryValueEx Lib "AdvApi32.dll" Alias "RegQueryValueExA" ( _
    ByVal HKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As Long, _
    ByRef LPType As Long, ByRef LPData As Any, ByRef lpcbData As Long) As Long

    Private Declare PtrSafe Function RegCloseKey Lib "AdvApi32.dll" (ByVal HKey As LongPtr) As LongPtr
    
    Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
    ByVal hWnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As LongPtr) As Long
#Else
    Private Declare Function SetParent Lib "user32" ( _
                             ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
                             
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
                             ByVal lpClassName As String, ByVal lpWindowName As String) As Long
                             
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
                             ByVal hwnd As Long, ByVal nIndex As Long) As Long
                             
    Private Declare Function EnableMenuItem Lib "user32" ( _
                             ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
                         
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
                             ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
                             
    Private Declare Function SetLayeredWindowAttributes Lib "user32" ( _
                             ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
                             
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
                             ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
                             
    Private Declare Function GetActiveWindow Lib "user32" () As Long
    
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
                             
    Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
                             
    Private Declare Function GetSystemMenu Lib "user32" ( _
                             ByVal hwnd As Long, ByVal bRevert As Long) As Long
                             
    Private Declare Function RemoveMenu Lib "user32" ( _
                             ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
                             
    Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
                             
    Private Declare Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
    
    Private Declare Function RegOpenKeyEx Lib "AdvApi32.dll" Alias "RegOpenKeyExA" ( _
                             ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, _
                             phkResult As Long) As Long
    
    Private Declare Function RegQueryValueEx Lib "AdvApi32.dll" Alias "RegQueryValueExA" ( _
                             ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
                             LPType As Long, LPData As Any, lpcbData As Long) As Long
    
    Private Declare Function RegCloseKey Lib "AdvApi32.dll" (ByVal HKey As LongPtr) As Long
    
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" ( _
                             ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
#End If

'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As FormControl
    Set Class = FormControl
End Property

Private Sub Class_Initialize()
    Guard.Singleton Me
End Sub

'@Description("Displays (if HideButton is False) or hides (if HideButton is True) a maximize window button _
NOTE: If EITHER a Minimize or Maximize button is displayed, BOTH buttons are visible but may be disabled.")
Public Function ShowMaximizeButton(ByRef outForm As MSForms.UserForm, ByVal HideButton As Boolean) As Boolean

    #If VBA7 Then
        Dim WinInfo As LongPtr
        Dim r As LongPtr
        Dim UFHWnd As LongPtr
    #Else
        Dim WinInfo As Long
        Dim r As Long
        Dim UFHWnd As Long
    #End If

    UFHWnd = HWndOfUserForm(outForm)
    If UFHWnd = 0 Then
        ShowMaximizeButton = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    If HideButton = False Then
        WinInfo = WinInfo Or WS_MAXIMIZEBOX
    Else
        WinInfo = WinInfo And (Not WS_MAXIMIZEBOX)
    End If
    r = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)

    ShowMaximizeButton = (r <> 0)
    
End Function

'@Description("Displays (if HideButton is False) or hides (if HideButton is True) a minimize window button. _
NOTE: If EITHER a Minimize or Maximize button is displayed, BOTH buttons are visible but may be disabled.")
Public Function ShowMinimizeButton(ByRef outForm As MSForms.UserForm, ByVal HideButton As Boolean) As Boolean
 
    #If VBA7 Then
        Dim WinInfo As LongPtr
        Dim r As LongPtr
        Dim UFHWnd As LongPtr
    #Else
        Dim WinInfo As Long
        Dim r As Long
        Dim UFHWnd As Long
    #End If

    UFHWnd = HWndOfUserForm(outForm)
    If UFHWnd = 0 Then
        ShowMinimizeButton = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    If HideButton = False Then
        WinInfo = WinInfo Or WS_MINIMIZEBOX
    Else
        WinInfo = WinInfo And (Not WS_MINIMIZEBOX)
    End If
    r = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)

    ShowMinimizeButton = (r <> 0)
    
End Function

'@Description("Returns True if the userform has a minimize button, False otherwise.")
Public Function HasMinimizeButton(ByRef outForm As MSForms.UserForm) As Boolean
        
    #If VBA7 Then
        Dim WinInfo As LongPtr
        Dim r As LongPtr
        Dim UFHWnd As LongPtr
    #Else
        Dim WinInfo As Long
        Dim r As Long
        Dim UFHWnd As Long
    #End If

    UFHWnd = HWndOfUserForm(outForm)
    If UFHWnd = 0 Then
        HasMinimizeButton = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

    If WinInfo And WS_MINIMIZEBOX Then
        HasMinimizeButton = True
    Else
        HasMinimizeButton = False
    End If
    
End Function

'@Description("Returns True if the userform has a maximize button, False otherwise.")
Public Function HasMaximizeButton(ByRef outForm As MSForms.UserForm) As Boolean
        
    #If VBA7 Then
        Dim WinInfo As LongPtr
        Dim r As LongPtr
        Dim UFHWnd As LongPtr
    #Else
        Dim WinInfo As Long
        Dim r As Long
        Dim UFHWnd As Long
    #End If

    UFHWnd = HWndOfUserForm(outForm)
    If UFHWnd = 0 Then
        HasMaximizeButton = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

    If WinInfo And WS_MAXIMIZEBOX Then
        HasMaximizeButton = True
    Else
        HasMaximizeButton = False
    End If
    
End Function

'@Description("Set the UserForm outForm as a child of (1) the Application, (2) the Excel ActiveWindow, or (3) no parent. Returns TRUE if successful or FALSE if unsuccessful.")
Public Function SetFormParent(ByRef outForm As MSForms.UserForm, ByVal Parent As FORM_PARENT_WINDOW_TYPE) As Boolean
        
    #If VBA7 Then
        Dim WindHWnd As LongPtr
        Dim r As LongPtr
        Dim UFHWnd As LongPtr
    #Else
        Dim WindHWnd As Long
        Dim r As Long
        Dim UFHWnd As Long
    #End If

    UFHWnd = HWndOfUserForm(outForm)
    If UFHWnd = 0 Then
        SetFormParent = False
        Exit Function
    End If

    Select Case Parent
    Case FORM_PARENT_APPLICATION
        r = SetParent(UFHWnd, Application.hWnd)
    Case FORM_PARENT_NONE
        r = SetParent(UFHWnd, 0&)
    Case FORM_PARENT_WINDOW
        If Application.ActiveWindow Is Nothing Then
            SetFormParent = False
            Exit Function
        End If
        WindHWnd = WindowHWnd(Application.ActiveWindow)
        If WindHWnd = 0 Then
            SetFormParent = False
            Exit Function
        End If
        r = SetParent(UFHWnd, WindHWnd)
    Case Else
        SetFormParent = False
        Exit Function
    End Select
    SetFormParent = (r <> 0)
    
End Function

'@Description("Returns TRUE if UserForm outForm has a close button, FALSE if there is no close button.")
Public Function IsCloseButtonVisible(ByRef outForm As MSForms.UserForm) As Boolean
        
    #If VBA7 Then
        Dim WinInfo As LongPtr
        Dim r As LongPtr
        Dim UFHWnd As LongPtr
    #Else
        Dim WinInfo As Long
        Dim r As Long
        Dim UFHWnd As Long
    #End If

    UFHWnd = HWndOfUserForm(outForm)
    If UFHWnd = 0 Then
        IsCloseButtonVisible = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    IsCloseButtonVisible = (WinInfo And WS_SYSMENU)
    
End Function

'@Description("This displays (if HideButton is FALSE) or hides (if HideButton is TRUE) the Close button on the userform")
Public Function ShowCloseButton(ByRef outForm As MSForms.UserForm, ByVal HideButton As Boolean) As Boolean

    #If VBA7 Then
        Dim WinInfo As LongPtr
        Dim r As LongPtr
        Dim UFHWnd As LongPtr
    #Else
        Dim WinInfo As Long
        Dim r As Long
        Dim UFHWnd As Long
    #End If

    UFHWnd = HWndOfUserForm(outForm)
    If UFHWnd = 0 Then
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    If HideButton = False Then
        ' set the SysMenu bit
        WinInfo = WinInfo Or WS_SYSMENU
    Else
        ' clear the SysMenu bit
        WinInfo = WinInfo And (Not WS_SYSMENU)
    End If

    r = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)
    ShowCloseButton = (r <> 0)
    
End Function

'@Description("This returns TRUE if the close button is enabled or FALSE if the close button is disabled.")
Public Function IsCloseButtonEnabled(ByRef outForm As MSForms.UserForm) As Boolean
        
    #If VBA7 Then
        Dim hMenu As LongPtr
        Dim ItemCount As LongPtr
        Dim PrevState As LongPtr
        Dim UFHWnd As LongPtr
    #Else
        Dim hMenu As Long
        Dim ItemCount As Long
        Dim PrevState As Long
        Dim UFHWnd As Long
    #End If

    UFHWnd = HWndOfUserForm(outForm)
    If UFHWnd = 0 Then
        IsCloseButtonEnabled = False
        Exit Function
    End If
    ' Get the menu handle
    hMenu = GetSystemMenu(UFHWnd, 0&)
    If hMenu = 0 Then
        IsCloseButtonEnabled = False
        Exit Function
    End If

    ItemCount = GetMenuItemCount(hMenu)
    ' Disable the button. This returns MF_DISABLED or MF_ENABLED indicating
    ' the previous state of the item.
    PrevState = EnableMenuItem(hMenu, ItemCount - 1, MF_DISABLED Or MF_BYPOSITION)

    If PrevState = MF_DISABLED Then
        IsCloseButtonEnabled = False
    Else
        IsCloseButtonEnabled = True
    End If
    ' restore the previous state
    EnableCloseButton outForm, (PrevState = MF_DISABLED)

    DrawMenuBar UFHWnd
        
End Function

'@Description("This function enables (if Disable is False) or disables (if Disable is True) the "X" button on a UserForm outForm.")
Public Function EnableCloseButton(ByRef outForm As MSForms.UserForm, ByVal Disable As Boolean) As Boolean
        
    #If VBA7 Then
        Dim hMenu As LongPtr
        Dim ItemCount As LongPtr
        Dim Res As LongPtr
        Dim UFHWnd As LongPtr
    #Else
        Dim hMenu As Long
        Dim ItemCount As Long
        Dim Res As Long
        Dim UFHWnd As Long
    #End If

    ' Get the HWnd of the UserForm.
    UFHWnd = HWndOfUserForm(outForm)
    If UFHWnd = 0 Then
        EnableCloseButton = False
        Exit Function
    End If
    ' Get the menu handle
    hMenu = GetSystemMenu(UFHWnd, 0&)
    If hMenu = 0 Then
        EnableCloseButton = False
        Exit Function
    End If

    ItemCount = GetMenuItemCount(hMenu)
    If Disable = True Then
        Res = EnableMenuItem(hMenu, ItemCount - 1, MF_DISABLED Or MF_BYPOSITION)
    Else
        Res = EnableMenuItem(hMenu, ItemCount - 1, MF_ENABLED Or MF_BYPOSITION)
    End If
    If Res = -1 Then
        EnableCloseButton = False
        Exit Function
    End If
    DrawMenuBar UFHWnd

    EnableCloseButton = True
    
End Function

'@Description("Displays (if HideTitle is FALSE) or hides (if HideTitle is TRUE) the title bar of the userform outForm.")
Public Function ShowTitleBar(ByRef outForm As MSForms.UserForm, ByVal HideTitle As Boolean) As Boolean
        
    #If VBA7 Then
        Dim WinInfo As LongPtr
        Dim r As LongPtr
        Dim UFHWnd As LongPtr
    #Else
        Dim WinInfo As Long
        Dim r As Long
        Dim UFHWnd As Long
    #End If

    UFHWnd = HWndOfUserForm(outForm)
    If UFHWnd = 0 Then
        ShowTitleBar = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

    If HideTitle = False Then
        ' turn on the Caption bit
        WinInfo = WinInfo Or WS_CAPTION
    Else
        ' turn off the Caption bit
        WinInfo = WinInfo And (Not WS_CAPTION)
    End If
    r = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)
    ShowTitleBar = (r <> 0)
    
End Function

'@Description("Returns TRUE if the title bar of outForm is visible or FALSE if the title bar is not visible.")
Public Function IsTitleBarVisible(ByRef outForm As MSForms.UserForm) As Boolean
        
    #If VBA7 Then
        Dim WinInfo As LongPtr
        Dim r As LongPtr
        Dim UFHWnd As LongPtr
    #Else
        Dim WinInfo As Long
        Dim r As Long
        Dim UFHWnd As Long
    #End If

    UFHWnd = HWndOfUserForm(outForm)
    If UFHWnd = 0 Then
        IsTitleBarVisible = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

    IsTitleBarVisible = (WinInfo And WS_CAPTION)
    
End Function

'@Description("This makes the userform outForm resizable (if Sizable is TRUE) or not resizable (if Sizalbe is FALSE). Returns TRUE if successful or FALSE if an error occurred.")
Public Function MakeFormResizable(ByRef outForm As MSForms.UserForm, ByVal Sizable As Boolean) As Boolean
        
    #If VBA7 Then
        Dim WinInfo As LongPtr
        Dim r As LongPtr
        Dim UFHWnd As LongPtr
    #Else
        Dim WinInfo As Long
        Dim r As Long
        Dim UFHWnd As Long
    #End If

    UFHWnd = HWndOfUserForm(outForm)
    If UFHWnd = 0 Then
        MakeFormResizable = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)
    If Sizable = True Then
        WinInfo = WinInfo Or WS_SIZEBOX
    Else
        WinInfo = WinInfo And (Not WS_SIZEBOX)
    End If

    r = SetWindowLong(UFHWnd, GWL_STYLE, WinInfo)
    MakeFormResizable = (r <> 0)

End Function

'@Description("Returns TRUE if outForm is resizable, FALSE if outForm is not resizable.")
Public Function IsFormResizable(ByRef outForm As MSForms.UserForm) As Boolean

    #If VBA7 Then
        Dim WinInfo As LongPtr
        Dim r As LongPtr
        Dim UFHWnd As LongPtr
    #Else
        Dim WinInfo As Long
        Dim r As Long
        Dim UFHWnd As Long
    #End If

    UFHWnd = HWndOfUserForm(outForm)
    If UFHWnd = 0 Then
        IsFormResizable = False
        Exit Function
    End If

    WinInfo = GetWindowLong(UFHWnd, GWL_STYLE)

    IsFormResizable = (WinInfo And WS_SIZEBOX)
    
End Function

'@Description("This function sets the opacity of the UserForm referenced by the outForm parameter. Opacity specifies the opacity of the form, from _
0 = fully transparent (invisible) to 255 = fully opaque. The function returns True if successful or False if an error occurred. This _
requires Windows 2000 or later -- it will not work in Windows 95, 98, or ME.")
Public Function SetFormOpacity(ByRef outForm As MSForms.UserForm, ByVal Opacity As Byte) As Boolean
        
    #If VBA7 Then
        Dim WinL As LongPtr
        Dim Res As LongPtr
        Dim UFHWnd As LongPtr
    #Else
        Dim WinL As Long
        Dim Res As Long
        Dim UFHWnd As Long
    #End If

    SetFormOpacity = False

    UFHWnd = HWndOfUserForm(outForm)
    If UFHWnd = 0 Then
        Exit Function
    End If

    WinL = GetWindowLong(UFHWnd, GWL_EXSTYLE)
    If WinL = 0 Then
        Exit Function
    End If

    Res = SetWindowLong(UFHWnd, GWL_EXSTYLE, WinL Or WS_EX_LAYERED)
    If Res = 0 Then
        Exit Function
    End If

    Res = SetLayeredWindowAttributes(UFHWnd, 0, Opacity, LWA_ALPHA)
    If Res = 0 Then
        Exit Function
    End If

    SetFormOpacity = True
    
End Function

'@Description("Clears the specified bit in Value and returns the result. Bits are _
numbered, right (most significant) 31 to left (least significant) 0. _
BitNumber is made positive and then MOD 32 to get a valid bit number.")
Public Function ClearBit(ByRef Value As Long, ByVal BitNumber As Long) As Long
    Dim SetMask As Long
    Dim ClearMask As Long

    BitNumber = Abs(BitNumber) Mod 32

    SetMask = Value
    If BitNumber < 30 Then
        ClearMask = Not (2 ^ (BitNumber - 1))
        ClearBit = SetMask And ClearMask
    Else
        ClearBit = Value And &H7FFFFFFF
    End If
End Function

'@Description("This returns the window handle (HWnd) of the userform referenced by outForm. _
It first looks for a top-level window, then a child of the Application window, then a child of the ActiveWindow.")
#If VBA7 Then
Public Function HWndOfUserForm(ByRef outForm As MSForms.UserForm) As LongPtr
#Else
Public Function HWndOfUserForm(ByRef outForm As MSForms.UserForm) As Long
#End If

#If VBA7 Then
    Dim UFHWnd As LongPtr
    Dim WinHWnd As LongPtr
#Else
    Dim UFHWnd As Long
    Dim WinHWnd As Long
#End If
    
    Dim AppHWnd As Long
    Dim DeskHWnd As Long
            
    Dim Cap As String
    Dim WindowCap As String
        
    Cap = outForm.Caption
        
    ' First, look in top level windows
    UFHWnd = FindWindow(C_USERFORM_CLASSNAME, Cap)
    If UFHWnd <> 0 Then
        HWndOfUserForm = UFHWnd
        Exit Function
    End If
    ' Not a top level window. Search for child of application.
    AppHWnd = Application.hWnd
    UFHWnd = FindWindowEx(AppHWnd, 0&, C_USERFORM_CLASSNAME, Cap)
    If UFHWnd <> 0 Then
        HWndOfUserForm = UFHWnd
        Exit Function
    End If

    ' Not a child of the application.
    ' Search for child of ActiveWindow (Excel's ActiveWindow, not
    ' Window's ActiveWindow).
    If Application.ActiveWindow Is Nothing Then
        HWndOfUserForm = 0
        Exit Function
    End If
    WinHWnd = WindowHWnd(Application.ActiveWindow)
    UFHWnd = FindWindowEx(WinHWnd, 0&, C_USERFORM_CLASSNAME, Cap)
    HWndOfUserForm = UFHWnd

End Function

' This function looks in the registry key HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced
' for the value named "HideFileExt" to determine whether the Windows Explorer setting "Hide Extensions Of Known File Types" is enabled. This function returns
' TRUE if this setting is in effect (meaning that Windows displays "Book1" rather than "Book1.xls"),
' or FALSE if this setting is not in effect (meaning that Windows displays "Book1.xls").
Private Function DoesWindowsHideFileExtensions() As Boolean

    #If VBA7 Then
        Dim Res As LongPtr
        Dim RegKey As LongPtr
    #Else
        Dim Res As Long
        Dim RegKey As Long
    #End If

    Dim v As Long

    Const KEY_NAME = "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Const VALUE_NAME = "HideFileExt"

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Open the registry key to get a handle (RegKey).
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    Res = RegOpenKeyEx(HKey:=HKCU, _
                       lpSubKey:=KEY_NAME, _
                       ulOptions:=0&, _
                       samDesired:=KEY_ALL_ACCESS, _
                       phkResult:=RegKey)

    If Res <> ERROR_SUCCESS Then
        Exit Function
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Get the value of the "HideFileExt" named value.
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    Res = RegQueryValueEx(HKey:=RegKey, _
                          lpValueName:=VALUE_NAME, _
                          lpReserved:=0&, _
                          LPType:=REG_DWORD, _
                          LPData:=v, _
                          lpcbData:=Len(v))

    If Res <> ERROR_SUCCESS Then
        RegCloseKey RegKey
        Exit Function
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Close the key and return the result.
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    RegCloseKey RegKey
    DoesWindowsHideFileExtensions = (v <> 0)

End Function

' This returns the Caption of the Excel.Window W with the ".xls" extension removed
' if required. The string returned by this function is suitable for use by
' the FindWindowEx API regardless of the value of the Windows "Hide Extensions" setting.
Private Function WindowCaption(ByRef W As Excel.Window) As String

    Dim HideExt As Boolean
    Dim Cap As String
    Dim Pos As Long

    HideExt = DoesWindowsHideFileExtensions()
    Cap = W.Caption
    If HideExt = True Then
        Pos = InStrRev(Cap, ".")
        If Pos > 0 Then
            Cap = VBA.Strings.Left(Cap, Pos - 1)
        End If
    End If

    WindowCaption = Cap

End Function

' This returns the HWnd of the Window referenced by W.
#If VBA7 Then
Private Function WindowHWnd(ByRef W As Excel.Window) As LongPtr
#Else
Private Function WindowHWnd(ByRef W As Excel.Window) As Long
#End If

#If VBA7 Then
    Dim DeskHWnd As LongPtr
    Dim WHWnd As LongPtr
#Else
    Dim DeskHWnd As Long
    Dim WHWnd As Long
#End If
    
    Dim AppHWnd As Long
    Dim Cap As String
        
    AppHWnd = Application.hWnd
    DeskHWnd = FindWindowEx(AppHWnd, 0&, C_EXCEL_DESK_CLASSNAME, vbNullString)
    If DeskHWnd > 0 Then
        Cap = WindowCaption(W)
        WHWnd = FindWindowEx(DeskHWnd, 0&, C_EXCEL_WINDOW_CLASSNAME, Cap)
    End If
    WindowHWnd = WHWnd

End Function

' This just wraps up GetWindowText.
Private Function WindowText(ByRef hWnd As Long) As String

    Dim s As String
    Dim N As Long
    N = 255
    s = VBA.String$(N, vbNullChar)
    N = GetWindowText(hWnd, s, N)
    If N > 0 Then
        WindowText = VBA.Strings.Left(s, N)
    Else
        WindowText = vbNullString
    End If
End Function

' This just wraps up GetClassName.
Private Function WindowClassName(ByRef hWnd As Long) As String

    Dim s As String
    Dim N As Long
    N = 255
    s = VBA.String$(N, vbNullChar)
    N = GetClassName(hWnd, s, N)
    If N > 0 Then
        WindowClassName = VBA.Strings.Left(s, N)
    Else
        WindowClassName = vbNullString
    End If

End Function


