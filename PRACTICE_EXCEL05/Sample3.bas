Attribute VB_Name = "Sample3"
Option Explicit

Sub HowMuch()
  Dim num1 As Single
  Dim num2 As Single
  Dim result As Single

  num1 = 45.33
  num2 = 19.24

  result = MultiplyIt(num1, num2)
  MsgBox result
End Sub


Function MultiplyIt(num1, num2) As Integer
  MultiplyIt = num1 * num2
End Function

