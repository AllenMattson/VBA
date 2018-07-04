Attribute VB_Name = "AllMarketData_ControlledbyMenu"
Option Explicit
Dim LR As Long, cell As Range
Sub Industries()
showsheet
Sheets("Market").Activate
Cells(2, 1).CurrentRegion.Copy

Workbooks.Add
ActiveSheet.Name = "Sheet1"
Range("A2").PasteSpecial xlPasteValues

Dim company As String, code As String
LR = Sheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row
Dim industryList As Range
Set industryList = Sheets("Sheet1").Range("A2:A" & LR)
'Set industryList = Range(Cells(=All_Sectors!$D$2:$D$102, 1), Cells(50, 1))
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
delsheets
ALLSECTORS
For Each cell In industryList
    cell.Offset(0, 1).Copy Sheet1.Range("B1")
    Sheets.Add
    ActiveSheet.Name = cell.value
    IND_URL_Static_Query (cell.Offset(0, 1).value)
Next


SortWorksheets
FIXB

Application.EnableEvents = True
KeepOtherSheetsHidden
ThisWorkbook.Saved = True
End Sub
Private Function IND_URL_Static_Query(cell)
Dim code As String
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
On Error Resume Next
code = Sheet1.Range("tag").value
ActiveSheet.Cells.Clear
   With ActiveSheet.QueryTables.Add(Connection:= _
      "URL;http://biz.yahoo.com/p/csv/" & code & "conameu.csv", _
         Destination:=Range("a1"))
      .BackgroundQuery = True
      .TablesOnlyFromHTML = True
      .Refresh BackgroundQuery:=False
      .SaveData = True
   End With
ActiveSheet.Range("a1").CurrentRegion.TextToColumns Destination:=ActiveSheet.Range("a1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=True, Space:=False, other:=False
ActiveSheet.Columns.AutoFit
End Function
Private Sub delsheets()
 Dim sh As Worksheet
Application.DisplayAlerts = False
For Each sh In ActiveWorkbook.Sheets
 If sh.Name <> "Sheet1" Then
    sh.Delete
 End If
Next sh
Application.DisplayAlerts = True
 End Sub

Sub ALLSECTORS()
Sheets.Add
ActiveSheet.Name = "All_Sectors"
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
ActiveSheet.Cells.Clear
   With ActiveSheet.QueryTables.Add(Connection:= _
      "URL;http://biz.yahoo.com/p/csv/s_conameu.csv", _
         Destination:=Range("a1"))
      .BackgroundQuery = True
      .TablesOnlyFromHTML = True
      .Refresh BackgroundQuery:=False
      .SaveData = True
   End With
ActiveSheet.Range("a1").CurrentRegion.TextToColumns Destination:=ActiveSheet.Range("a1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=True, Space:=False, other:=False
ActiveSheet.Columns.AutoFit
End Sub
Private Sub SortWorksheets()
     
    Dim N As Integer
    Dim M As Integer
    Dim FirstWSToSort As Integer
    Dim LastWSToSort As Integer
    Dim SortDescending As Boolean
     
    SortDescending = False
     
    If ActiveWindow.SelectedSheets.Count = 1 Then
         
         'Change the 1 to the worksheet you want sorted first
        FirstWSToSort = 1
        LastWSToSort = Worksheets.Count
    Else
        With ActiveWindow.SelectedSheets
            For N = 2 To .Count
                If .Item(N - 1).Index <> .Item(N).Index - 1 Then
                    MsgBox "You cannot sort non-adjacent sheets"
                    Exit Sub
                End If
            Next N
            FirstWSToSort = .Item(1).Index
            LastWSToSort = .Item(.Count).Index
        End With
    End If
     
    For M = FirstWSToSort To LastWSToSort
        For N = M To LastWSToSort
            If SortDescending = True Then
                If UCase(Worksheets(N).Name) > UCase(Worksheets(M).Name) Then
                    Worksheets(N).Move Before:=Worksheets(M)
                End If
            Else
                If UCase(Worksheets(N).Name) < UCase(Worksheets(M).Name) Then
                    Worksheets(N).Move Before:=Worksheets(M)
                End If
            End If
        Next N
    Next M
     
End Sub
Sub FIXB()
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    LR = Cells(Rows.Count, 10).End(xlUp).Row
    Dim RNG1 As Range
    Set RNG1 = Range("K2:K" & LR)
    Dim cell As Range
    For Each cell In RNG1
        If Not IsEmpty(cell.value) Then
            cell.Offset(0, -10).value = cell.Offset(0, -10).value & cell.Offset(0, -9).value
            cell.Offset(0, -9).Delete shift:=xlToLeft
        End If
    Next cell
Next ws
End Sub
Private Sub Workbook_SheetSelectionChange(ByVal sh As Object, ByVal Target As Range)

Dim ans As Integer
ans = MsgBox("Do You Want to Navigate to a New Sheet?", vbYesNo)
Select Case ans
    Case vbYes
        UserForm.Show vbModeless
    Case vbNo
        ActiveSheet.Select
End Select
End Sub
Private Sub KeepOtherSheetsHidden()
Application.ScreenUpdating = False
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    If ws.Visible = True And ws.Name <> "All_Sectors" Then
        ws.Visible = xlSheetHidden
    End If
Next
Sheets("All_Sectors").Select
Application.ScreenUpdating = False
End Sub
Sub showsheet()
Sheet1.Visible = True
End Sub
Sub ProfitChart()
Dim Udata As Range
ActiveSheet.Range("a5").Select
Set Udata = ActiveCell.CurrentRegion


Dim LR As Long, RNG As Range, RNG2 As Range, RNG3 As Range
LR = ActiveSheet.UsedRange.Rows.Count
Dim ChObj As ChartObject
For Each ChObj In ActiveSheet.ChartObjects
    ChObj.Delete
Next

'Set RNG = "='Sector Services'!$E$2:$E$62"
Set RNG = Range(Cells(2, 5), Cells(LR, 5))
Set RNG2 = Range(Cells(2, 9), Cells(LR, 9))

ActiveSheet.Shapes.AddChart2(240, xlXYScatter).Select
On Error Resume Next
With ActiveChart
    .SetSourceData Source:=Udata
    .FullSeriesCollection(1).Delete
    .FullSeriesCollection(6).Delete
    .FullSeriesCollection(5).Delete
    .FullSeriesCollection(4).Delete
    .FullSeriesCollection(3).Delete
    .FullSeriesCollection(2).Delete
    .FullSeriesCollection(1).Delete
    .SeriesCollection.NewSeries
    .FullSeriesCollection(1).XValues = RNG
    .FullSeriesCollection(1).Values = RNG2
    .HasTitle = False
    .HasLegend = False
End With
AttachLabelsToPoints
MENU.Show vbModeless
End Sub
Sub AttachLabelsToPoints()

   'Dimension variables.
   Dim Counter As Integer, ChartName As String, xVals As String

   ' Disable screen updating while the subroutine is run.
   Application.ScreenUpdating = False

   'Store the formula for the first series in "xVals".
   xVals = ActiveChart.SeriesCollection(1).Formula

   'Extract the range for the data from xVals.
   xVals = Mid(xVals, InStr(InStr(xVals, ","), xVals, _
      Mid(Left(xVals, InStr(xVals, "!") - 1), 9)))
   xVals = Left(xVals, InStr(InStr(xVals, "!"), xVals, ",") - 1)
   Do While Left(xVals, 1) = ","
      xVals = Mid(xVals, 2)
   Loop

   'Attach a label to each data point in the chart.
   For Counter = 1 To Range(xVals).Cells.Count
     ActiveChart.SeriesCollection(1).Points(Counter).HasDataLabel = _
         True
      ActiveChart.SeriesCollection(1).Points(Counter).DataLabel.Text = _
         Range(xVals).Cells(Counter, 1).Offset(0, -4).value
   Next Counter

End Sub



