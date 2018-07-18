Attribute VB_Name = "Sample8"
Option Explicit

Sub Informant()
  InputBox prompt:="Enter your place of birth:" & Chr(13) _
      & " (e.g., Boston, Great Falls, etc.) "
End Sub


Sub Informant2()
  Dim myPrompt As String
  Dim town As String

  Const myTitle = "Enter data"

  myPrompt = "Enter your place of birth:" & Chr(13) _
      & "(e.g., Boston, Great Falls, etc.)"
  town = InputBox(myPrompt, myTitle)

  MsgBox "You were born in " & town & ".", , "Your response"
End Sub

Sub AddTwoNums()
  Dim myPrompt As String
  Dim value1 As String
  Dim value2 As Integer
  Dim mySum As Single

  Const myTitle = "Enter data"

  myPrompt = "Enter a number:"
  value1 = InputBox(myPrompt, myTitle, 0)
  value2 = 2
  mySum = value1 + value2

  MsgBox "The result is " & mySum & _
       " (" & value1 & " + " & CStr(value2) + ")", _
        vbInformation, "Total"
End Sub

Sub WhatRange()
  Dim newRange As Range
  Dim tellMe As String

  tellMe = "Use the mouse to select a range:"
  Set newRange = Application.InputBox(prompt:=tellMe, _
      Title:="Range to format", _
      Type:=8)
  newRange.NumberFormat = "0.00"
  newRange.Select
End Sub

Sub WhatRange2()
  Dim newRange As Range
  Dim tellMe As String

  On Error GoTo VeryEnd

  tellMe = "Use the mouse to select a range:"
  Set newRange = Application.InputBox(prompt:=tellMe, _
      Title:="Range to format", _
      Type:=8)
  newRange.NumberFormat = "0.00"
  newRange.Select

VeryEnd:
End Sub

