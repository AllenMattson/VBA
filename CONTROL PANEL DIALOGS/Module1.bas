Attribute VB_Name = "Module1"
Sub RunSysProg()
    UserForm1.Show
End Sub

Sub ShowDateTimeDlg()
  Arg = "rundll32.exe shell32.dll,Control_RunDLL timedate.cpl"
  On Error Resume Next
  TaskID = Shell(Arg)
  If Err <> 0 Then
      MsgBox ("Cannot start the application.")
  End If
End Sub

