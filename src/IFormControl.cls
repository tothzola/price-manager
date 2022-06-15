VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFormControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("FormControl")
'@Exposed
'@Interface

Option Explicit

Public Enum FORM_PARENT_WINDOW_TYPE
    FORM_PARENT_NONE = 0
    FORM_PARENT_APPLICATION = 1
    FORM_PARENT_WINDOW = 2
End Enum

'Displays or hides a Maximize Window button on the userform.
Public Function ShowMaximizeButton(UF As MSForms.UserForm, HideButton As Boolean) As Boolean
End Function

'Displays or hides a Minimize Window button on the userform.
Public Function ShowMinimizeButton(UF As MSForms.UserForm, HideButton As Boolean) As Boolean
End Function

'Returns True or False indicating whether the userform has a minimize button.
Public Function HasMinimizeButton(UF As MSForms.UserForm) As Boolean
End Function

'Returns True or False indicating whether the userform has a maximize button.
Public Function HasMaximizeButton(UF As MSForms.UserForm) As Boolean
End Function

'Sets a userform's parent to the Application or the ActiveWindow.
Public Function SetFormParent(UF As MSForms.UserForm, Parent As FORM_PARENT_WINDOW_TYPE) As Boolean
End Function

'Returns True or False indicating whether the userform's Close button is visible.
Public Function IsCloseButtonVisible(UF As MSForms.UserForm) As Boolean
End Function

'Displays or hides the userform's Close button.
Public Function ShowCloseButton(UF As MSForms.UserForm, HideButton As Boolean) As Boolean
End Function

'Returns True or False indicating whether the userform's Close button is enabled.
Public Function IsCloseButtonEnabled(UF As MSForms.UserForm) As Boolean
End Function

'Enables or disables a userform's Close button.
Public Function EnableCloseButton(UF As MSForms.UserForm, Disable As Boolean) As Boolean
End Function

'Displays or hides a userform's Title Bar. The title bar cannot be hidden if the form is resizable.
Public Function ShowTitleBar(UF As MSForms.UserForm, HideTitle As Boolean) As Boolean
End Function

'Returns True or False indicating if the userform's Title Bar is visible.
Public Function IsTitleBarVisible(UF As MSForms.UserForm) As Boolean
End Function

'Makes the form resizable or not resizable. If the form is made resizable, the title bar cannot be hidden.
Public Function MakeFormResizable(UF As MSForms.UserForm, Sizable As Boolean) As Boolean
End Function

'Returns True or False indicating whether the userform is resizable.
Public Function IsFormResizable(UF As MSForms.UserForm) As Boolean
End Function

'Sets the opacity of a form from fully opaque to fully invisible.
Public Function SetFormOpacity(UF As MSForms.UserForm, Opacity As Byte) As Boolean
End Function

'Returns the window handle (HWnd) of a userform.
#If VBA7 Then
Private Function HWndOfUserForm(UF As MSForms.UserForm) As LongPtr
#Else
Private Function HWndOfUserForm(UF As MSForms.UserForm) As Long
#End If
End Function

'Clears the specified bit in Value and returns the result. Bits are numbered, right (most significant) 31 to left (least significant) 0.
Public Function ClearBit(Value As Long, ByVal BitNumber As Long) As Long
End Function

