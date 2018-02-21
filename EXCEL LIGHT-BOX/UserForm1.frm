VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Background"
   ClientHeight    =   555
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   1770
   OleObjectBlob   =   "UserForm1.frx":0000
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function GetActiveWindow Lib "user32.dll" () As Long
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function BringWindowToTop Lib "user32.dll" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function BringWindowToTop Lib "user32.dll" (ByVal hWnd As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
#End If


Const GWL_STYLE = -16
Const WS_CAPTION = &HC00000
Const WS_SYSMENU = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2
Dim hWnd As Long

Private Sub UserForm_Initialize()
    Dim lngWindow As Long, lFrmHdl As Long
    lFrmHdl = FindWindow(vbNullString, Me.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
End Sub

Private Sub UserForm_activate()
    Dim ufcap As String
    ufcap = UserForm1.Caption
    hWnd = FindWindow("ThunderDFrame", ufcap)

'   Adjust UserForm to Excel's window size
    With Me
        .Height = Application.Height
        .Width = Application.Width
        .Left = Application.Left
        .Top = Application.Top
    End With
    TransparentUserForm Me, 180 'increase to make darker
    Select Case Application.Caller
        Case "Button 1"
            Call ShowPicture
        Case "Button 2"
            Call ShowMsgBox
    End Select
End Sub

Private Function TransparentUserForm(frm As UserForm, Level As Byte) As Boolean
'   Makes a UserForm transparent, semi-transparent, or invisible
'   Level: 0 to 255
    SetWindowLong hWnd, GWL_EXSTYLE, WS_EX_LAYERED
    SetLayeredWindowAttributes hWnd, 0, Level, LWA_ALPHA
    TranslucentForm = Err.LastDllError = 0
End Function

Sub ShowPicture()
'   Displays on top of semi-transparent UserForm
    With UserForm2
        'make sure it's centered
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
      .Show
    End With
    Unload Me
End Sub

Sub ShowMsgBox()
    Dim Msg As String
    Msg = "This is just a demonstration of how to get a 'light-box' effect "
    Msg = Msg & "in Excel. The window is covered with a semitransparent black UserForm, "
    Msg = Msg & "and the message box is displayed on top."
    MsgBox Msg, vbInformation, "Light-Box Demo"
    Unload UserForm1
End Sub
