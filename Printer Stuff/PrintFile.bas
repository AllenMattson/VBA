Attribute VB_Name = "PrintFile"
Option Explicit
 
Declare Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long
 
Public Sub PrintFile(ByVal strPathAndFilename As String)
 
    Call apiShellExecute(Application.hwnd, "print", strPathAndFilename, vbNullString, vbNullString, 0)
 
End Sub
 
Sub Test()
'Application.CommandBars.ExecuteMso ("PrintPreviewAndPrint")
'Application.Dialogs(xlDialogPrint).Show
PrintFile ("C:\Test.pdf")

End Sub

