Attribute VB_Name = "Sample5"
Option Explicit

Function Avg(num1, num2, Optional num3)
  Dim totalNums As Integer

  totalNums = 3

  If IsMissing(num3) Then
    num3 = 0
    totalNums = totalNums - 1
  End If

  Avg = (num1 + num2 + num3) / totalNums
End Function

