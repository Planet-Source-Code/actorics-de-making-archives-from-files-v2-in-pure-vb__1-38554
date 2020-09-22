Attribute VB_Name = "modFlat"
Option Explicit
'Some APIs and Constants

'We need this API because it get's our frmMain-Window (or any control on it)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
'This API we need for to set some Window (Control) attributes
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'This API we nned for to if we want to resize it or if we want to move it
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal CX As Long, ByVal CY As Long, ByVal wFlags As Long) As Long
'Here are the constants for the APIs
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4

Sub MakeFlat(lHwnd As Long)
'Adds a Flatlook to any Object
'Of course if a window like frmMain's borderstyle is = 3 then we can't make the form flat
    Dim lRet As Long
    'Get the hWnd of a control or window
    lRet = GetWindowLong(lHwnd, GWL_EXSTYLE)
    'Set lRet-Flags
    lRet = lRet Or WS_EX_STATICEDGE And Not WS_EX_CLIENTEDGE
    'Set the FlatBorder
    SetWindowLong lHwnd, GWL_EXSTYLE, lRet
    'Set some WindowPropertys
    SetWindowPos lHwnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_FRAMECHANGED
End Sub
