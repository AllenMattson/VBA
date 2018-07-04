Attribute VB_Name = "OpenNotePad"
Option Explicit

Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, _
ByVal hWndChildAfter As Long, ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) _
As Long

Sub test()
Dim hwnd&, TaskIdChk&, TaskID&

TaskID = Shell("C:\Windows\notepad.exe", 1)

hwnd = FindWindowEx(0, 0, "notepad", vbNullString)
GetWindowThreadProcessId hwnd, TaskIdChk

If TaskIdChk <> TaskID Then 'Just to check if found handle is the correct one
 MsgBox "The handle is not for the created task!"
 Exit Sub
End If

Do Until IsWindow(hwnd) = 0
Loop

End Sub
