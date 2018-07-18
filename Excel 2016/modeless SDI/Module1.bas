Attribute VB_Name = "Module1"
Option Explicit

Dim WA As New WinActivate

Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" _
  (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

Public UserFormHandle As Long

Sub ShowModeless()
    Set WA.AppEvents = Application
    UserForm1.Show 0
    UserFormHandle = FindWindow("ThunderDFrame", UserForm1.Caption)
End Sub


