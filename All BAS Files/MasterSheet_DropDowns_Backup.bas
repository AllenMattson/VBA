Attribute VB_Name = "MasterSheet_DropDowns"
Option Explicit
Sub Main_Master()

Dim LR As Integer
Dim wb As Workbook: Set wb = ThisWorkbook
'Make sure correct workbook is selected
If wb.Name <> "Master.xlsm" Then
Dim msg As String
    msg = MsgBox("Please activate the Master Workbook", vbCritical, "Wrong Workbook")
    Exit Sub
End If


Sheets("Sheet1").Select

Dim nLastCol As Long, LastRo As Long

'counting variables for loops
Dim i As Integer, j As Integer, k As Integer, t As Integer


'delete named ranges if any
Dim sName As Name
For Each sName In Names
    sName.Delete
Next


Sheets("Sheet2").Activate

nLastCol = Cells(1, Columns.Count).End(xlToLeft).Column 'Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column



Sheets("Sheet2").Activate

Dim myRANGE As Range, MyList As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Create Named Ranges to Build Dynamic Drop Down Lists''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RNGstr As String
Dim N As Name
For i = 1 To nLastCol
'Sheets("Sheet2").Activate

    If Cells(1, i) <> "" Then
        Sheets("Sheet2").Activate
        LastRo = Sheets("Sheet2").Cells(Rows.Count, i).End(xlUp).Row
        Set myRANGE = ActiveSheet.Range(Cells(2, i), Cells(LastRo, i))
        MyList = Cells(1, i).Text
        ActiveSheet.Range(Cells(2, i), Cells(LastRo, i)).Name = MyList
    End If
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Re-enter the list drop downs
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'delete drop downs in case columns have been changed/updated
Sheets("Sheet1").Activate
With ActiveSheet
With ActiveSheet.Cells.Validation
    .Delete
End With

'Rebuild
With Range("B2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=topic"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
End With

With Range("C2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=lead_source"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
End With
  
With Range("D2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=status"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
End With

With Range("E2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=program"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
End With

With Range("F2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=bus_adviser"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
End With

With Range("AV2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=refferal_forwards"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
End With
Range("AV2").AutoFill Destination:=Range("AV2:AV900"), Type:=xlFillDefault
Range("AV2:AV900").Select: Selection.Interior.Color = vbYellow

Range("B2").AutoFill Destination:=Range("B2:B900"), Type:=xlFillDefault
Range("B2:B900").Select: Selection.Interior.Color = vbYellow

Range("C2").AutoFill Destination:=Range("C2:C900"), Type:=xlFillDefault
Range("C2:C900").Select: Selection.Interior.Color = vbYellow

Range("D2").AutoFill Destination:=Range("D2:D900"), Type:=xlFillDefault
Range("D2:D900").Select: Selection.Interior.Color = vbYellow

Range("E2").AutoFill Destination:=Range("E2:E900"), Type:=xlFillDefault
Range("E2:E900").Select: Selection.Interior.Color = vbYellow

Range("F2").AutoFill Destination:=Range("F2:F900"), Type:=xlFillDefault
Range("F2:F900").Select: Selection.Interior.Color = vbYellow

End With
Columns.AutoFit
Rows.AutoFit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets("Sheet2").Activate: Cells(1, 1).Select
End Sub
