Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 And Win64 Then
    Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#Else
    Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
#End If

'   Constants for the keys of interest
    Const VK_SHIFT As Integer = &H10
    Const VK_CONTROL As Integer = &H11
    Const VK_MENU As Integer = &H12 'Alt key


Sub DisplayKeyStatus()
    Dim TabChar As String * 1
    Dim CRChar As String * 1
    Dim Shift As Boolean, Control As Boolean, Alt As Boolean
    Dim Msg As String
    
    TabChar = Chr(9)
    CRChar = Chr(13)

'   Use API calls to determine which keys are pressed
    If GetKeyState(VK_SHIFT) < 0 Then Shift = True Else Shift = False
    If GetKeyState(VK_CONTROL) < 0 Then Control = True Else Control = False
    If GetKeyState(VK_MENU) < 0 Then Alt = True Else Alt = False

'   Build the message
    Msg = "Shift:" & TabChar & Shift & CRChar
    Msg = Msg & "Control:" & TabChar & Control & CRChar
    Msg = Msg & "Alt:" & TabChar & Alt & CRChar
    
'   Display message box
    MsgBox Msg, vbInformation, "Key Status"
End Sub

