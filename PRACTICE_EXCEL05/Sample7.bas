Attribute VB_Name = "Sample7"
Option Explicit

Sub MsgYesNo()
  Dim question As String
  Dim myButtons As Integer

  question = "Do you want to open a new workbook?"
  myButtons = vbYesNo + vbQuestion + vbDefaultButton2

  MsgBox question, myButtons
End Sub

Sub MsgYesNo3()
  Dim question As String
  Dim myButtons As Integer
  Dim myTitle As String

  Dim myChoice As Integer

  question = "Do you want to open a new workbook?"
  myButtons = vbYesNo + vbQuestion + vbDefaultButton2
  myTitle = "New workbook"
  myChoice = MsgBox(question, myButtons, myTitle)

  MsgBox myChoice
End Sub

