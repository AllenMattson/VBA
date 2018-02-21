Attribute VB_Name = "Module1"
Option Explicit

Sub Add100Buttons()
  Dim UFvbc As VBComponent
  Dim CMod As CodeModule
  Dim ctl As Control
  Dim cb As CommandButton
  Dim n As Long, c As Long, r As Long
  Dim code As String
  
'   Make sure access to the VBProject is allowed
    On Error Resume Next
    Dim x
    Set x = ActiveWorkbook.VBProject
    If Err <> 0 Then
        MsgBox "Your security settings do not allow this macro to run.", vbCritical
        On Error GoTo 0
        Exit Sub
    End If
  
  Set UFvbc = ThisWorkbook.VBProject.VBComponents("UserForm1")

' Delete all controls, if any
  For Each ctl In UFvbc.Designer.Controls
    UFvbc.Designer.Controls.Remove ctl.Name
  Next ctl

' Delete all VBA code
  UFvbc.CodeModule.DeleteLines 1, UFvbc.CodeModule.CountOfLines

' Add 100 CommandButtons
  n = 1
  For r = 1 To 10
    For c = 1 To 10
      Set cb = UFvbc.Designer.Controls.Add("Forms.CommandButton.1")
      With cb
        .Width = 22
        .Height = 22
        .Left = (c * 26) - 16
        .Top = (r * 26) - 16
        .Caption = n
      End With

'     Add the event handler code
      With UFvbc.CodeModule
        code = ""
        code = code & "Private Sub CommandButton" & n & "_Click" & vbCr
        code = code & "Msgbox ""This is CommandButton" & n & """" & vbCr
        code = code & "End Sub"
        .InsertLines .CountOfLines + 1, code
      End With
      n = n + 1
    Next c
  Next r
  VBA.UserForms.Add("UserForm1").Show
End Sub

Sub ShowForm()
  UserForm1.Show
End Sub

