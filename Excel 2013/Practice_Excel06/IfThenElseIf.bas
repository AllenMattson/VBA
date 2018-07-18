Attribute VB_Name = "IfThenElseIf"
Option Explicit

Sub WhatValue()
  Range("A1").Select
  If ActiveCell.Value = 0 Then
    ActiveCell.Offset(0, 1).Value = "zero"
  ElseIf ActiveCell.Value > 0 Then
    ActiveCell.Offset(0, 1).Value = "positive"
  ElseIf ActiveCell.Value < 0 Then
    ActiveCell.Offset(0, 1).Value = "negative"
  End If
End Sub

