Attribute VB_Name = "modWorksheet"
Option Explicit

Sub WorksheetSort(list)
'   Sorts an array by transferring it to
'   A worksheet and using Excel's sorting command
    
    Dim First As Integer, Last As Long
    Dim i As Long
    Dim FirstCell As Range, LastCell As Range
    Dim CurrCell As Range, FillRange As Range
    
    First = LBound(list, 1)
    Last = UBound(list, 1)
    Set FirstCell = Sheets("Sheet1").Cells(1, 1)
    Set LastCell = Sheets("Sheet1").Cells(Last, 1)
    Set FillRange = Range(FirstCell, LastCell)
    Application.ScreenUpdating = False
    
'   Transfer the array to worksheet
    FillRange.Value = list
    
'   Sort the worksheet range
    FirstCell.CurrentRegion.Sort Key1:=FirstCell, Order1:=xlAscending, Orientation:=xlTopToBottom
       
'   Transfer range back to the array and clear range
    For i = First To Last
        list(i, 0) = FirstCell.Offset(i - 1, 0)
    Next i
    FillRange.Clear
    Application.ScreenUpdating = True
End Sub

