VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CapsLockClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' API declarations
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetVersionEx Lib "Kernel32" _
        Alias "GetVersionExA" _
        (lpVersionInformation As OSVERSIONINFO) As Long
    
    Private Declare PtrSafe Sub keybd_event Lib "user32" _
        (ByVal bVk As Byte, _
        ByVal bScan As Byte, _
        ByVal dwflags As Long, ByVal dwExtraInfo As Long)
    
    Private Declare PtrSafe Function GetKeyboardState Lib "user32" _
        (pbKeyState As Byte) As Long
    
    Private Declare PtrSafe Function SetKeyboardState Lib "user32" _
        (lppbKeyState As Byte) As Long
#Else
    Private Declare Function GetVersionEx Lib "Kernel32" _
        Alias "GetVersionExA" _
        (lpVersionInformation As OSVERSIONINFO) As Long
    
    Private Declare Sub keybd_event Lib "user32" _
        (ByVal bVk As Byte, _
        ByVal bScan As Byte, _
        ByVal dwflags As Long, ByVal dwExtraInfo As Long)
    
    Private Declare Function GetKeyboardState Lib "user32" _
        (pbKeyState As Byte) As Long
    
    Private Declare Function SetKeyboardState Lib "user32" _
        (lppbKeyState As Byte) As Long
#End If


' Type declaration
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type


'Constant declarations
Const VK_NUMLOCK = &H90
Const VK_SCROLL = &H91
Const VK_CAPITAL = &H14
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
      
Property Get Value() As Boolean
    Dim keys(0 To 255) As Byte
    GetKeyboardState keys(0)
    Value = keys(VK_CAPITAL)
End Property

Property Let Value(boolVal As Boolean)
    Dim o As OSVERSIONINFO
    Dim keys(0 To 255) As Byte
    o.dwOSVersionInfoSize = Len(o)
    GetVersionEx o
    GetKeyboardState keys(0)
'   Is it already in that state?
    If boolVal = True And keys(VK_CAPITAL) = 1 Then Exit Property
    If boolVal = False And keys(VK_CAPITAL) = 0 Then Exit Property
'   Toggle it
    'Simulate Key Press
    keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
    'Simulate Key Release
    keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
End Property

Sub Toggle()
'   Toggles the state
    Dim o As OSVERSIONINFO
    Dim keys(0 To 255) As Byte
    o.dwOSVersionInfoSize = Len(o)
    GetVersionEx o
    GetKeyboardState keys(0)
    'Simulate Key Press
    keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
    'Simulate Key Release
    keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
End Sub

