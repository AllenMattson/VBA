VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function GetKeyboardState Lib "user32" _
    (pbKeyState As Byte) As Long

Private Declare PtrSafe Function SetKeyboardState Lib "user32" _
    (lppbKeyState As Byte) As Long
    
Private Declare Sub keybd_event Lib "user32" _
    (ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    
Const VK_NUMLOCK = &H90
Const VK_SCROLL = &H91
Const VK_CAPITAL = &H14
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2

Public Property Get NumLock() As Boolean
    Dim Keys(0 To 255) As Byte
    GetKeyboardState Keys(0)
    NumLock = Keys(VK_NUMLOCK)
End Property

Public Property Get CapsLock() As Boolean
    Dim Keys(0 To 255) As Byte
    GetKeyboardState Keys(0)
    CapsLock = Keys(VK_CAPITAL)
End Property

Public Property Get ScrollLock() As Boolean
    Dim Keys(0 To 255) As Byte
    GetKeyboardState Keys(0)
    ScrollLock = Keys(VK_SCROLL)
End Property

Public Sub ToggleNumLock()
    'Simulate Key Press
    keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
    'Simulate Key Release
    keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY _
        Or KEYEVENTF_KEYUP, 0
End Sub

Public Sub ToggleCapsLock()
    'Simulate Key Press
      keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
    'Simulate Key Release
      keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY _
         Or KEYEVENTF_KEYUP, 0
End Sub

Public Sub ToggleScrollLock()
    'Simulate Key Press
      keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
    'Simulate Key Release
      keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY _
        Or KEYEVENTF_KEYUP, 0
End Sub
