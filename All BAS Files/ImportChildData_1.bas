Attribute VB_Name = "ImportChildData_1"
Sub MergeFilesWithoutSpacesSTART()
    Dim path As String, ThisWB As String, lngFilecounter As Long
    Dim wbDest As Workbook, shtDest As Worksheet, ws As Worksheet
    Dim FileName As String, Wkb As Workbook
    Dim CopyRng As Range, Dest As Range
    Dim RowofCopySheet As Integer
ThisWB = ActiveWorkbook.Name
'Sheets("Sheet3").Activate
'path = Sheets("Sheet3").Cells(1, 2).Value
Sheets("Sheet1").Activate
Dim i As Integer, lr As Integer: lr = Cells(Rows.Count, 1).End(xlUp).Row
For i = lr To 2 Step -1
    Rows(i).Delete
Next i
Sheets("sheet3").Activate
If Range("B1").Value <> "" Then
    path = Sheets("Sheet3").Cells(1, 2).Value
Else
    path = InputBox("Please Insert Folder Path to Child Sheets", "BA Folder Location", ThisWorkbook.path)
    Sheets("Sheet3").Cells(1, 2).Value = path
End If
If path = "" Then
    MsgBox "Insert Path to Import BA sheets"
    Exit Sub
End If






    If Right(path, 1) <> "\" Then
        If Right(path, 1) <> ":" Then

            If IsMac = True Then
                Cells(1, 2).Clear
                If Right(path, 1) <> ":" Then path = path & ":"
                Cells(1, 2).Value = path
                Cells(3, 2).Clear
                Cells(3, 2).Value = "Mac"
            
            Else
                If IsMac = False Then
                    Cells(1, 2).Clear
                    If Right(path, 1) <> "\" Then path = path & "\"
                    Cells(1, 2).Value = path
                    Cells(3, 2).Clear
                    Cells(3, 2).Value = "PC"
                End If
            End If
        End If
    End If
Sheets("Sheet1").Activate
RowofCopySheet = 2

Application.EnableEvents = False
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Set shtDest = ActiveWorkbook.Sheets(1) '("test")
FileName = Dir(path & "*.xlsx", vbNormal)
If Len(FileName) = 0 Then Exit Sub
Do Until FileName = vbNullString
    If Not FileName = ThisWB Then
    Application.DisplayAlerts = False
        Set Wkb = Workbooks.Open(FileName:=path & FileName)
        Set CopyRng = Wkb.Sheets(1).Cells(1, 1).CurrentRegion 'Range(Cells(RowofCopySheet, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, Cells(1, Columns.Count).End(xlToLeft).Column))
        Set Dest = shtDest.Range("A" & shtDest.Cells(Rows.Count, 1).End(xlUp).Row + 1)
        CopyRng.Copy
        Dest.PasteSpecial xlPasteFormats
        Dest.PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False 'Clear Clipboard
        Wkb.Close False
    End If
FileName = Dir()
Loop

Dim k As Integer, LaRo As Integer
LaRo = Cells(Rows.Count, "T").End(xlUp).Row
For k = LaRo To 2 Step -1
    If Trim(Cells(k, 1).Value) = "Date" Then Rows(k).Delete
    If Trim(Cells(k, 1).Value) = "MyDate" Then Rows(k).Delete
Next k

Dim Nlr As Long: Nlr = Cells(Rows.Count, "T").End(xlUp).Row
Range("A2:ZZ" & Nlr).Copy
Range("A2:ZZ" & Nlr).PasteSpecial xlPasteValues
Range("A2:ZZ" & Nlr).Validation.Delete
Cells(1, 1).Select

    

Application.EnableEvents = True
Application.DisplayAlerts = True
Application.ScreenUpdating = True

Dim M As Workbook: Set M = ActiveWorkbook
M.Save
End Sub

Function IsMac() As Boolean
#If Mac Then
    IsMac = True
#ElseIf Win32 Or Win64 Then
    IsMac = False
#End If
End Function
