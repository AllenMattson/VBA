Attribute VB_Name = "ForEachNextLoop"
Option Explicit

Sub RemoveSheets()
  Dim mySheet As Worksheet

  Application.DisplayAlerts = False

  Workbooks.Add
  Sheets.Add After:=ActiveSheet, Count:=3

  For Each mySheet In Worksheets
    If mySheet.Name <> "Sheet1" Then
        ActiveWindow.SelectedSheets.Delete
    End If
  Next mySheet

  Application.DisplayAlerts = True
End Sub

Sub IsSuchSheet(strSheetName As String)
  Dim mySheet As Worksheet
  Dim counter As Integer

  counter = 0

  Workbooks.Add
  Sheets.Add After:=ActiveSheet, Count:=3
    For Each mySheet In Worksheets
        If mySheet.Name = strSheetName Then
            counter = counter + 1
            Exit For
        End If
    Next mySheet

    If counter = 1 Then
        MsgBox strSheetName & " exists."
    Else
        MsgBox strSheetName & " was not found."
    End If
End Sub


Sub FindSheet()
   Call IsSuchSheet("Sheet2")
End Sub

Sub EarlyExit()
  Dim myCell As Variant
  Dim myRange As Range
    
  Set myRange = Range("A1:H10")
  For Each myCell In myRange
      If myCell.Value = "" Then
          myCell.Value = "empty"
      Else
          Exit For
      End If
  Next myCell
End Sub

Sub ColorLoop()
  Dim myRow As Integer
  Dim myCol As Integer
  Dim myColor As Integer

  myColor = 0

  For myRow = 1 To 8
    For myCol = 1 To 7
        Cells(myRow, myCol).Select
        myColor = myColor + 1
        With Selection.Interior
          .ColorIndex = myColor
          .Pattern = xlSolid
        End With
    Next myCol
  Next myRow
End Sub


