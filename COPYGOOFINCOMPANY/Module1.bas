Attribute VB_Name = "Module1"
Option Explicit
Sub addEmptyColumn()
Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
If ws.Visible = True Then
    ws.Activate
    MoveIt
End If
Next

End Sub
Sub test()
Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    If ws.Visible = True And ws.Name <> "Sheet1" Then
        ws.Delete
    End If
Next
Application.DisplayAlerts = True
End Sub
Sub AddCol()
Dim lc As Long
lc = Cells(1, Columns.Count).End(xlToLeft).Column
Dim i As Integer
For i = 2 To Columns.Count
    ActiveSheet.Columns(i).Insert Shift:=xlToRight
    i = i + 1
Next

End Sub
Sub ALLINDUSTRIES()


    Dim j As Integer

    On Error Resume Next

    Worksheets.Add ' add a sheet in first place
    Sheets(1).Name = "All Companies"

    ' copy headings
    Sheets(1).Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets(1).Range("A1")

    ' work through sheets
    For j = 1 To Sheets.Count ' from sheet 2 to last sheet
        Sheets(j).Activate ' make the sheet active
        Range("A1").Select
        Selection.CurrentRegion.Select ' select all cells in this sheets

        ' select all lines except title
        Selection.Offset(1, 0).Resize(Selection.Rows.Count).Select

        ' copy cells selected in the new sheet on last line
        Selection.Copy Destination:=Sheets(1).Range("A65536").End(xlUp)(2)
    Next
ActiveSheet.Range("A1").Select

End Sub
Sub MoveIt()

Dim LastRow As Long
Dim ws1 As Worksheet

Set ws1 = ActiveSheet

Do While (ws1.Range("B1").Value <> "")
    LastRow = ws1.Range("A" & ws1.Rows.Count).End(xlUp).Row + 1
    ws1.Range("B1:B" & ws1.Range("B" & ws1.Rows.Count).End(xlUp).Row).Copy
    ws1.Range("A" & LastRow).PasteSpecial
    ws1.Range("B1").EntireColumn.Delete xlToLeft
Loop

End Sub
Sub DeleteBlanks()
    Dim intCol As Integer
     
    For intCol = 1 To 1 'cols A to D
        Range(Cells(1, intCol), Cells(8184, intCol)). _
        SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
    Next intCol
End Sub
