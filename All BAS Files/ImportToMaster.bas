Attribute VB_Name = "ImportToMaster"
Sub MergeFilesWithoutSpaces()
    Dim path As String, ThisWB As String, lngFilecounter As Long
    Dim wbDest As Workbook, shtDest As Worksheet, ws As Worksheet
    Dim Filename As String, Wkb As Workbook
    Dim CopyRng As Range, Dest As Range
    Dim RowofCopySheet As Integer
ThisWB = ActiveWorkbook.Name

path = Sheets("Sheet3").Cells(1, 2).Value

RowofCopySheet = 2

Application.EnableEvents = False
Application.ScreenUpdating = False

Set shtDest = ActiveWorkbook.Sheets(1) '("test")
Filename = Dir(path & "\*.xlsx", vbNormal)
If Len(Filename) = 0 Then Exit Sub
Do Until Filename = vbNullString
    If Not Filename = ThisWB Then
    Application.DisplayAlerts = False
        Set Wkb = Workbooks.Open(Filename:=path & "\" & Filename)
        Set CopyRng = Wkb.Sheets(1).Range(Cells(RowofCopySheet, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, Cells(1, Columns.Count).End(xlToLeft).Column))
        Set Dest = shtDest.Range("A" & shtDest.Cells(Rows.Count, 1).End(xlUp).Row + 1)
        CopyRng.Copy
        Dest.PasteSpecial xlPasteFormats
        Dest.PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False 'Clear Clipboard
        Wkb.Close False
    End If

    Filename = Dir()
Loop

Dim k As Integer, LaRo As Integer
LaRo = Cells(Rows.Count, 1).End(xlUp).Row
For k = LaRo To 2 Step -1
    If Cells(k, 1).Value = "Date" Then Rows(k).Delete
Next k
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

