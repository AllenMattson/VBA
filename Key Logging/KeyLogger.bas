Attribute VB_Name = "KeyLogger"
Option Explicit

Declare Function SetWindowsHookEx Lib _
"user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, _
ByVal hmod As Long, ByVal dwThreadId As Long) As Long

Declare Function CallNextHookEx Lib "user32" _
(ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Declare Function GetActiveWindow Lib "user32" () As Long

Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Const HC_ACTION = 0
Const WM_KEYDOWN = &H100
Const WH_KEYBOARD_LL = 13
Dim hhkLowLevelKybd As Long
Dim blnHookEnabled As Boolean
Dim enumAllowedValues As AllowedValues
Dim objTargetRange As Range
Dim objValidationRange As Range
Dim vAns As Variant

Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Enum AllowedValues
    alpha
    numeric
End Enum




Function LowLevelKeyboardProc _
(ByVal nCode As Long, ByVal wParam As Long, lParam As KBDLLHOOKSTRUCT) As Long

    '\hook keyboard only if XL is the active window
    If GetActiveWindow = FindWindow("XLMAIN", Application.Caption) Then
        If (nCode = HC_ACTION) Then
            '\check if a key is pushed
            If wParam = WM_KEYDOWN Then
            '\if so, check if the active cell is within the target range
                If Union(ActiveCell, objTargetRange).Address = objTargetRange.Address Then
                '\if only numeric values should be allowed then
                    If enumAllowedValues = 1 Then
                    '\check if the pushed key is a numeric key or a navigation key
                    '\by checking the vkCode stored in the laparm parameter
                        If Chr(lParam.vkCode) Like "#" Or _
                            lParam.vkCode = 37 Or lParam.vkCode = 38 Or lParam.vkCode = 39 Or _
                            lParam.vkCode = 40 Or lParam.vkCode = 9 Or lParam.vkCode = 13 Then
                            '\if so allow the input
                            LowLevelKeyboardProc = 0
                        Else
                            '\else filter out this Key_Down message from message qeue
                            Beep
                            LowLevelKeyboardProc = -1
                            Exit Function
                        End If
                        '\if onle alpha values should be allowed then
                    ElseIf enumAllowedValues = 0 Then
                        '\check the laparam parameter
                        If Chr(lParam.vkCode) Like "#" Then
                            '\if numeric prevent the input
                            Beep
                            LowLevelKeyboardProc = -1
                            Exit Function
                        Else
                            '\otherwise allow the input
                            LowLevelKeyboardProc = 0
                    End If
                    End If
                End If
            End If
        End If
    End If
    '\pass function to next hook if there is one
    LowLevelKeyboardProc = CallNextHookEx(0, nCode, wParam, ByVal lParam)

End Function


Public Sub Unhook_KeyBoard()

    If hhkLowLevelKybd <> 0 Then UnhookWindowsHookEx hhkLowLevelKybd
    blnHookEnabled = False
    Cells.Clear

End Sub


Sub ValidateRange(r As Range, ByVal v As AllowedValues)

    '\store these in global variables for they will be
    '\needed later in the filter function
    enumAllowedValues = v
    Set objTargetRange = r
    '\don't hook the keyboard twice !!
    If blnHookEnabled = False Then
        hhkLowLevelKybd = SetWindowsHookEx _
        (WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, Application.Hinstance, 0)
        blnHookEnabled = True
    End If

End Sub


Sub test()

    '\ignore any mishandling of the following
    '\input boxes by the user
    On Error Resume Next
    Cells.Clear
    Set objValidationRange = Application.InputBox _
    ("Selet one or more Cells ", "Custom Data Validation...", Type:=8)
    If objValidationRange Is Nothing Then GoTo errHdlr
        objValidationRange.Interior.Color = vbGreen
        vAns = InputBox("To allow only alpha values in the selected range enter 1 " _
        & vbCrLf & vbCrLf & "To allow only numeric values in the selected range enter 2 ")
        If vAns = 1 Then
            ValidateRange objValidationRange, AllowedValues.alpha
        ElseIf vAns = 2 Then
            ValidateRange objValidationRange, AllowedValues.numeric
        Else
        GoTo errHdlr
    End If
    objValidationRange.Cells(1).Select
    Set objValidationRange = Nothing
    Exit Sub
errHdlr:
    MsgBox "criteria error- Try again !", vbCritical
    Unhook_KeyBoard

End Sub

