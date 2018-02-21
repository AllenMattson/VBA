Attribute VB_Name = "Module1"
Option Explicit


' API FUNCTIONS DECLARATIONS

Declare Function FindWindow Lib "user32" Alias _
    "FindWindowA" (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Declare Function SendMessageA Lib "user32" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Integer, ByVal lParam As Long) As Long

Declare Function ExtractIconA Lib "shell32.dll" _
    (ByVal hInst As Long, ByVal lpszExeFileName As String, _
    ByVal nIconIndex As Long) As Long
    
Declare Function GetActiveWindow Lib "user32.dll" () As Long

Declare Function SetWindowPos Lib "user32" _
                    (ByVal hWnd As Long, _
                     ByVal hWndInsertAfter As Long, _
                     ByVal x As Long, _
                     ByVal Y As Long, _
                     ByVal cx As Long, _
                     ByVal cy As Long, _
                     ByVal wFlags As Long) As Long
                                       
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Declare Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal hWnd As Long, ByVal crey As Byte, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

                                       
' variable declarations
Public hWnd As Long

' Constant declarations
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_STYLE = (-16)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_APPWINDOW = &H40000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_THICKFRAME = &H40000
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40
Public Const SW_SHOW = 5
Public Const HWND_TOP = 0
Public Const LWA_ALPHA = &H2&
Public Const WM_SETICON = &H80




