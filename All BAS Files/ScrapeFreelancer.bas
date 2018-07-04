Attribute VB_Name = "ScrapeFreelancer"
Sub HTML_Table_To_Excel()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False
Dim sh As Worksheet

If ActiveSheet.Name <> "Freelancer Jobs" Then: Sheets.Add: ActiveSheet.Name = "Freelancer Jobs"
For Each sh In ThisWorkbook.Worksheets
    If sh.Name <> "Freelancer Jobs" Then sh.Delete
Next sh

Cells.ClearContents
Columns.AutoFit: Rows.AutoFit

    Dim htm As Object
    Dim Tr As Object
    Dim Td As Object
    Dim Tab1 As Object
 'https://www.freelancer.com/jobs/j-fixed/2/?cl=l-en
 Dim k As Integer
 For k = 1 To 500
    'Replace the URL of the webpage that you want to download
    Web_URL = "https://www.freelancer.com/jobs/j-fixed/" & k & "?cl=l-en" 'VBA.Trim(Sheets(1).Cells(1, 1))
 
    'Create HTMLFile Object
    Set HTML_Content = CreateObject("htmlfile")
 
    'Get the WebPage Content to HTMLFile Object
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", Web_URL, False
        .send
        HTML_Content.Body.Innerhtml = .responseText
    End With
 
    Column_Num_To_Start = 1
    irow = 2
    iCol = Column_Num_To_Start
    iTable = 0
    If Sheets(1).Cells(irow, iCol).Value <> "" Then
        irow = Cells(Rows.Count, 1).End(xlUp).Row
    End If
    'Loop Through Each Table and Download it to Excel in Proper Format
    For Each Tab1 In HTML_Content.getElementsByTagName("table")
        With HTML_Content.getElementsByTagName("table")(iTable)
            For Each Tr In .Rows
                For Each Td In Tr.Cells
                On Error GoTo NextData
                    If Td.innertext = "Project/Contest " Or Len(Td.innertext) < 3 Then GoTo NextData
                    Sheets(1).Cells(irow, iCol).Select
                    Sheets(1).Cells(irow, iCol) = Td.innertext
                    iCol = iCol + 1
                Next Td
NextData:
                iCol = Column_Num_To_Start
                irow = irow + 1
            Next Tr
        End With
        iTable = iTable + 1
        iCol = Column_Num_To_Start
        irow = irow + 1
    Next Tab1
Next k
Columns("A:A").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
'Cells.RowHeight = 65: Cells.HorizontalAlignment = xlLeft: Cells.VerticalAlignment = xlCenter
'Cells.WrapText = True

With Cells(1, 1).CurrentRegion
    .ColumnWidth = 44
    .CurrentRegion.Borders.LineStyle = xlContinuous
    .CurrentRegion.Borders.Weight = xlMedium
    .RowHeight = 65
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlCenter
    .WrapText = True
End With
Cells(1, 1).Value = "PROJECT/CONTEST"
Cells(1, 2).Value = "DESCRIPTION"
Cells(1, 3).Value = "BIDS"
Cells(1, 4).Value = "KEYWORDS"
Cells(1, 5).Value = "DATE POSTED"
Cells(1, 6).Value = "TIME POSTED"
Cells(1, 7).Value = "PRICE"

With ActiveSheet
    .Sort.SortFields.Clear
    .Sort.SortFields.Add Key:=Range("G1"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With Worksheets("Freelancer Jobs").Sort
        .SetRange Cells(1, 1).CurrentRegion
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End With


Cells(1, 1).AutoFilter

Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.DisplayAlerts = True
MsgBox "Process Completed"
ActiveSheet.ListObjects.Add _
(xlSrcRange, Cells(1, 1).CurrentRegion, , xlYes).Name = "Freelancer_Jobs"
End Sub

