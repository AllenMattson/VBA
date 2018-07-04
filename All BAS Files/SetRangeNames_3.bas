Attribute VB_Name = "SetRangeNames_3"
Sub ListsRemakeStart()
Dim PassWrd As String, MSG As String
Dim word As String: word = "overlord"
PassWrd = InputBox("Please enter Password to run this macro", "Run New Drop Down Lists and Recreate Master Table")

If PassWrd = word Then
    SetRangeNames_3.PermissionGranted
Else
    MsgBox "YOU NEED A PASSWORD FOR THIS MACRO TO RUN!" & vbNewLine & vbNewLine & "Enter Password Correctly to Build New Table Lists", vbOKOnly, "Correct Password Needed..."
    Exit Sub
End If

End Sub
Sub PermissionGranted()
Dim wb As Workbook: Set wb = ActiveWorkbook
Application.ScreenUpdating = False
Dim Msh As Worksheet: Set Msh = Sheets(1)
Dim sh As Worksheet: Set sh = Sheets(2)
Dim LR As Integer, LRR As Integer, LC As Integer, LCC As Integer
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

DoEvents
With wb
sh.Activate
LC = sh.Cells(1, Columns.Count).End(xlToLeft).Column


For j = 1 To LC
sh.Activate
    If Trim(sh.Cells(1, j)) <> "" Then
        
        LR = sh.Cells(Rows.Count, j).End(xlUp).Row
        MyStr = sh.Cells(1, j).Value
        MyStr = Trim(MyStr)
        'create named range
        ActiveSheet.Range(Cells(2, j), Cells(LR, j)).Name = "List_" & j
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
                        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=("=List_" & j)
                        .IgnoreBlank = True
                        .InCellDropdown = True
                        .ShowInput = True
                        .ShowError = True
                    End With
                    NextEmpty.Select
                    'Fill Down Validation Formula, Clear Any Formats
                    Selection.AutoFill Destination:=Range(Cells(LRR + 1, i), Cells(900, i)), Type:=xlFillDefault
                    Range(Cells(LRR, 1).Offset(1, 0), Cells(900, i)).Select
                    Range(Cells(LRR + 1, i), Cells(900, i)).Select
                    Range(Cells(1, i), Cells(LRR, i)).ClearFormats
            
            End If
        Next i
    End If
    'activate the list sheet again for next range named
    sh.Activate
Next j
Msh.Activate

Application.ScreenUpdating = True
Columns.AutoFit
End With
End Sub
