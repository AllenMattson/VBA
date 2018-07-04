Attribute VB_Name = "SetRangeNames_1"
Sub ListsRemakeStart()

'delete old sheets if there is any
Dim OldSH As Worksheet
For Each OldSH In ThisWorkbook.Sheets
    If OldSH.Name <> "Sheet1" Then
        If OldSH.Name <> "Sheet2" Then
            If OldSH.Name <> "Sheet3" Then
                Application.DisplayAlerts = False
                OldSH.Delete
                Application.DisplayAlerts = True
            End If
        End If
    End If
Next OldSH
Application.ScreenUpdating = False
Dim Msh As Worksheet: Set Msh = Sheets("Sheet1")
Dim sh As Worksheet: Set sh = Sheets("Sheet2")
Dim lr As Integer, LRR As Integer, LC As Integer, LCC As Integer
Dim i As Integer, j As Integer
Dim MyStr As String
Dim NextEmpty As Range
'Loop through lists and create named ranges
'Enter in a list box on sheet one with the name of the list


'delete old range names
Msh.Activate
Dim n As Name
For Each n In ActiveSheet.Names
    If n <> "" Then n.Delete
Next n
'delete old validation cells
With ActiveSheet.Cells.Validation
    .Delete
End With
Range("A2:ZZ1000").Cells.Clear
Range(Cells(2, 1), Cells(90000, 90)).Cells.Interior.ColorIndex = xlNone
DoEvents
'Cells.Clear: sh.Rows("1:1").Copy
'Msh.Rows("1:1").Insert: Application.CutCopyMode = False
sh.Activate
LC = sh.Cells(1, Columns.Count).End(xlToLeft).Column


For j = 1 To LC
sh.Activate
    If Trim(sh.Cells(1, j)) <> "" Then
        
        lr = sh.Cells(Rows.Count, j).End(xlUp).Row
        MyStr = sh.Cells(1, j).Value
        MyStr = Trim(MyStr)
        'create named range
        ActiveSheet.Range(Cells(2, j), Cells(lr, j)).Name = "List_" & j
        Cells(1, 1).Select
        'find the header column in master sheet
        'place named range into next empty cell, fill down
        Msh.Activate: Cells(1, 1).Select
        LCC = Msh.Cells(1, Columns.Count).End(xlToLeft).Column
        'Loop to locate header
        For i = 1 To LCC
            If Trim(Cells(1, i).Value) = MyStr Then
                LRR = Cells(Rows.Count, i).End(xlUp).Row
                Range(Cells(LRR, i), Cells(LRR, i)).Select
                Set NextEmpty = ActiveCell.Offset(1, 0)
                
                NextEmpty.Select
            '   Enter in list validation
                    With Selection.Validation
                        .Delete
                        .add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=("=List_" & j)
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .ShowInput = True
                        .ShowError = True
                    End With
                    NextEmpty.Select
                    Selection.AutoFill Destination:=Range(Cells(LRR + 1, i), Cells(900, i)), Type:=xlFillDefault
                    Range(Cells(LRR, 1).Offset(1, 0), Cells(900, i)).Select
                    Range(Cells(LRR + 1, i), Cells(900, i)).Interior.Color = vbYellow
                    Range(Cells(LRR + 1, i), Cells(900, i)).Select
            '           remove old formats
                    Range(Cells(1, i), Cells(LRR, i)).ClearFormats
            
            End If
        Next i
    End If
    'activate the list sheet again for next range named
    sh.Activate
Next j
Msh.Activate
Range("T2:AW900").Interior.Color = vbGreen
Range("A2:B900").Interior.Color = vbYellow: Range("H2:S900").Interior.Color = vbYellow: Range("AX2:BB900").Interior.Color = vbYellow
Application.ScreenUpdating = True
Columns.AutoFit
Sheets("Sheet3").Activate: Range("B1:B6").Cells.Clear: Cells(1, 1).Select
End Sub
