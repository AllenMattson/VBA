Attribute VB_Name = "DoLoops"
Option Explicit

Sub ApplyBold()
  Do While ActiveCell.Value <> ""
    ActiveCell.Font.Bold = True
    ActiveCell.Offset(1, 0).Select
  Loop
End Sub

Sub TenSeconds()
  Dim stopme

  stopme = Now + TimeValue("00:00:10")

  Do While Now < stopme
    Application.DisplayStatusBar = True
    Application.StatusBar = Now
  Loop

  Application.StatusBar = False
End Sub

Sub SignIn()
      Dim secretCode As String
      Do
        secretCode = InputBox("Enter your secret code:")
        If secretCode = "sp1045" Then Exit Do
      Loop While secretCode <> "sp1045"
End Sub


Sub SayHello()
'press Ctrl+Break to stop this infinite loop
  Do
    MsgBox "Hello."
  Loop
End Sub

Sub ApplyBold2()
      Do Until IsEmpty(ActiveCell)
        ActiveCell.Font.Bold = True
        ActiveCell.Offset(1, 0).Select
      Loop
End Sub

Sub DeleteBlankSheets()
  Dim myRange As Range
  Dim shcount As Integer
  shcount = Worksheets.Count
  Do
    Worksheets(shcount).Select
    Set myRange = ActiveSheet.UsedRange
    If myRange.Address = "$A$1" And _
        Range("A1").Value = "" Then
        Application.DisplayAlerts = False
        Worksheets(shcount).Delete
        Application.DisplayAlerts = True
    End If
    shcount = shcount - 1
  Loop Until shcount = 1
End Sub



