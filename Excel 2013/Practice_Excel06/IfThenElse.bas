Attribute VB_Name = "IfThenElse"
Option Explicit

Sub WhatTypeOfDay()
  Dim response As String
  Dim question As String
  Dim strmsg1 As String, strmsg2 As String
  Dim myDate As Date

  question = "Enter any date in the format mm/dd/yyyy:" _
    & Chr(13) & " (e.g., 11/22/2013 )"
    strmsg1 = "weekday"
    strmsg2 = "weekend"
    response = InputBox(question)
    myDate = Weekday(CDate(response))
    If myDate >= 2 And myDate <= 6 Then
      MsgBox strmsg1
    Else
      MsgBox strmsg2
    End If
End Sub


Sub EnterData()
  Dim cell As Object
  Dim strmsg As String

  On Error GoTo VeryEnd

  strmsg = "Select any cell:"
  Set cell = Application.InputBox(prompt:=strmsg, Type:=8)
  cell.Select

  If IsEmpty(ActiveCell) Then
    ActiveCell.Formula = InputBox("Enter text or number:")
  Else
    ActiveCell.Offset(1, 0).Select
  End If

VeryEnd:
End Sub


