Attribute VB_Name = "Shell_Function"
Option Explicit

Sub StartPanel()
    Shell "Control.exe", vbNormalFocus
End Sub

Sub ChangeSettings()
   Dim nrTask
   nrTask = Shell("Control.exe intl.cpl", vbMinimizedFocus)
   Debug.Print nrTask
End Sub

Sub ChangeSettings2()
    Dim nrTask
    Dim iconFile As String
    iconFile = InputBox("Enter the name of the control " & _
              "icon CPL or DLL file:")
    nrTask = Shell("Control.exe " & iconFile, vbMinimizedFocus)
    Debug.Print nrTask
End Sub

Sub RunWord()
    Dim ReturnValue As Variant
    ReturnValue = Shell("C:\Program Files\Microsoft Office 15\" & _
        "root\office15\WINWORD.EXE", 1)
    ' activate Microsoft Word
    AppActivate ReturnValue
End Sub




