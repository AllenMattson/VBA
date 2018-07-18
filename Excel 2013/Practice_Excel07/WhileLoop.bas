Attribute VB_Name = "WhileLoop"
Option Explicit

Sub ChangeRHeight()
  While ActiveCell <> ""
     ActiveCell.RowHeight = 28
     ActiveCell.Offset(1, 0).Select
  Wend
End Sub

