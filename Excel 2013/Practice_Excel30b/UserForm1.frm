VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4665
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Activate()
    CustomizeForm
End Sub

Private Sub UserForm_Initialize()
    With Me
        .Caption = "Customized Form"
        .BackColor = RGB(255, 255, 51)
    End With
    AddIcon_OnTitleBar "C:\Excel2013_HandsOn\Extra Images\arrow.bmp"
End Sub


Private Sub AddIcon_OnTitleBar(strIconBmpFile As String)
    Dim fLen As Long
    
    If Len(Dir(strIconBmpFile)) <> 0 Then
        fLen = ExtractIconA(0, strIconBmpFile, 0)
        SendMessageA FindWindow(vbNullString, Me.Caption), _
            WM_SETICON, False, fLen
    Else
        Exit Sub
    End If
End Sub

Private Sub CustomizeForm()
    Dim wStyle As Long
    Dim xStyle As Long
    Dim bOpacity As Byte

        
    'get the handle of the active window
    hWnd = GetActiveWindow
    
    bOpacity = 150 ' set opacity
    
    'retrieve the active window's styles
    wStyle = GetWindowLong(hWnd, GWL_STYLE)
    
    ' modify the window style settings
    wStyle = wStyle Or WS_MINIMIZEBOX   'add the minimize button
    wStyle = wStyle Or WS_MAXIMIZEBOX   'add the maximize button
    wStyle = wStyle Or WS_THICKFRAME    'add a sizing border

    'apply the revised style
    Call SetWindowLong(hWnd, GWL_STYLE, wStyle)
             
    'retrieve the active window's extended styles
    xStyle = GetWindowLong(hWnd, GWL_EXSTYLE)

    ' modify the window extended style settings
    xStyle = xStyle Or WS_EX_LAYERED    ' change opacity
    xStyle = xStyle Or WS_EX_APPWINDOW  ' add window to the task bar

    'apply the revised extended style
    Call SetWindowLong(hWnd, GWL_EXSTYLE, xStyle)
    
    Call SetLayeredWindowAttributes(hWnd, 0, bOpacity, LWA_ALPHA)
    
    Call SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, _
                          SWP_NOMOVE Or _
                          SWP_NOSIZE Or _
                          SWP_NOACTIVATE Or _
                          SWP_HIDEWINDOW)

    Call SetWindowPos(hWnd, HWND_TOP, 0, 0, 0, 0, _
                          SWP_NOMOVE Or _
                          SWP_NOSIZE Or _
                          SWP_NOACTIVATE Or _
                          SWP_SHOWWINDOW)

End Sub




