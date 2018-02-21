VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNumLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Constant declarations
Const VK_NUMLOCK = &H90
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2

Private Declare Sub keybd_event Lib "user32" _
    (ByVal bVk As Byte, _
    ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
          
Private Declare PtrSafe Function GetKeyboardState Lib "user32" _
    (pbKeyState As Byte) As Long

Private Declare PtrSafe Function SetKeyboardState Lib "user32" _
    (lppbKeyState As Byte) As Long

Public Property Get Value() As Boolean
'   Get the current state
    Dim Keys(0 To 255) As Byte
    GetKeyboardState Keys(0)
    Value = CBool(Keys(VK_NUMLOCK))
End Property

Public Sub Toggle()
'   Toggles the state
    'Simulate Key Press
    keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
    'Simulate Key Release
    keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY _
        Or KEYEVENTF_KEYUP, 0
End Sub
