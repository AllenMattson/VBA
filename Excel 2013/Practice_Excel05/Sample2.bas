Attribute VB_Name = "Sample2"
Option Explicit

Sub EnterText()
  Dim m As String, n As String, r As String

  m = InputBox("Enter your first name:")
  n = InputBox("Enter your last name:")
  r = JoinText(m, n)

  MsgBox r
End Sub

Function JoinText(k, o)
  JoinText = k + " " + o
End Function

Function NumOfDays()
  NumOfDays = 7
End Function

Sub DaysInAWeek()
  MsgBox "There are " & NumOfDays & " days in a week."
End Sub

