Attribute VB_Name = "IfThen"
Option Explicit

Sub SimpleIfThen()
  Dim weeks As String
  On Error GoTo VeryEnd
  weeks = InputBox("How many weeks are in a year:", "Quiz")
  If weeks <> 52 Then MsgBox "Try Again": SimpleIfThen
  If weeks = 52 Then MsgBox "Congratulations!"
VeryEnd:
End Sub


Sub IfThenAnd()
  Dim price As Single
  Dim units As Integer
  Dim rebate As Single

  Const strmsg1 = "To get a rebate you must buy an additional "
  Const strmsg2 = "Price must equal $7.00"

  units = Range("B1").Value
  price = Range("B2").Value

  If price = 7 And units >= 50 Then
    rebate = (price * units) * 0.1
    Range("A4").Value = "The rebate is: $" & rebate
  End If

  If price = 7 And units < 50 Then
    Range("A4").Value = strmsg1 & 50 - units & " unit(s)."
  End If

  If price <> 7 And units >= 50 Then
    Range("A4").Value = strmsg2
  End If

  If price <> 7 And units < 50 Then
    Range("A4").Value = "You didn't meet the criteria."
  End If
End Sub

