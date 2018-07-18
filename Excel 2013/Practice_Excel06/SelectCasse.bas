Attribute VB_Name = "SelectCasse"
Option Explicit

Sub TestButtons()
  Dim question As String
  Dim bts As Integer
  Dim myTitle As String
  Dim myButton As Integer

  question = "Do you want to open a new workbook?"
  bts = vbYesNoCancel + vbQuestion + vbDefaultButton1
  myTitle = "New Workbook"
  myButton = MsgBox(prompt:=question, _
      Buttons:=bts, _
      Title:=myTitle)
  Select Case myButton
    Case 6
          Workbooks.Add
    Case 7
          MsgBox "You can open a new book manually later."
        Case Else
          MsgBox "You pressed Cancel."
  End Select
End Sub


Sub DisplayDiscount()
  Dim unitsSold As Integer
  Dim myDiscount As Single
  unitsSold = InputBox("Enter the number of units sold:")
  myDiscount = GetDiscount(unitsSold)
  MsgBox myDiscount
End Sub


Function GetDiscount(unitsSold As Integer)
  Select Case unitsSold
    Case 1 To 200
      GetDiscount = 0.05
    Case Is <= 500
      GetDiscount = 0.1
    Case 501 To 1000
      GetDiscount = 0.15
    Case Is > 1000
      GetDiscount = 0.2
  End Select
End Function


