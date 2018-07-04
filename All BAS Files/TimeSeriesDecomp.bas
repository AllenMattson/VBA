Attribute VB_Name = "TimeSeriesDecomp"
Sub Start_Main()
Sheets("Sheet1").Activate
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'*********************REQUEST USER INPUT BOXES B1:B4***************************'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Add Asset Drop Downs
Dim LC As Integer: LC = Range("A7").End(xlToRight).Column
Dim LR As Integer: LR = Cells(Rows.Count, 1).End(xlUp).Row



Dim Arng As Range: Set Arng = Range("B7", Cells(7, LC))
Application.CutCopyMode = False
ActiveWorkbook.Names.Add Name:="assets", RefersTo:=Arng
Sheets("TimeSeries").Select
With Range("C2").Validation
    .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=("=assets")
    .IgnoreBlank = True
    .InCellDropdown = True
    .ShowInput = True
    .ShowError = True
End With

End Sub
Sub MovingAverageValues()
Application.CutCopyMode = False
If ActiveSheet.Name <> "TimeSeries" Then Sheets("TimeSeries").Activate
If Range("C2").Value = "" Then
    MsgBox "Missing Asset"
    Exit Sub
Else
    If Range("C4").Value = "" Then
        MsgBox "Missing Moving Average"
        Exit Sub
    End If
End If
    
Range("J5:K900000").Cells.ClearContents
Range("F5:G9").Cells.ClearContents
Dim Str As String: Str = Range("C2").Value
Sheets("Sheet1").Activate
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'*********************REQUEST USER INPUT BOXES B1:B4***************************'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Add Asset Drop Downs
Dim LC As Integer: LC = Range("A7").End(xlToRight).Column
Dim LR As Integer: LR = Cells(Rows.Count, 1).End(xlUp).Row
Dim DateRNG As Range: Set DateRNG = Range(Cells(8, 1), Cells(LR, 1))
Dim ValRNG As Range
Dim CombinedRNG As Range

Dim Arng As Range: Set Arng = Range("B7", Cells(7, LC))
Dim i As Integer
For i = 1 To LC
    If Cells(7, i) = Str Then
       Set ValRNG = Range(Cells(8, i), Cells(LR, i))
            With ValRNG
                Cells.Replace "#N/A", "0", xlPart
            End With
    End If
Next i
Set CombinedRNG = Union(DateRNG, ValRNG)
CombinedRNG.Copy Sheets("TimeSeries").Range("J5")
If ActiveSheet.Name <> "TimeSeries" Then Sheets("TimeSeries").Activate
Range("F5:G9").Select: Selection.FormulaArray = "=LINEST(R5C11:R6201C11,R5C10:R6201C10,TRUE,TRUE)"
Range("A1").Select
End Sub
