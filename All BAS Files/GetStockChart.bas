Attribute VB_Name = "GetStockChart"
Option Explicit

Sub GetData()
    Dim DataSheet As Worksheet
    Dim EndDate As Date
    Dim StartDate As Date
    Dim Symbol As String
    Dim qurl As String
    Dim nQuery As Name
    Dim LastRow As Integer
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    Sheets("Data").Cells.Clear
    
    Set DataSheet = Sheets("CandleChart")
  
        StartDate = DataSheet.Range("startDate").value
        EndDate = DataSheet.Range("endDate").value
        Symbol = DataSheet.Range("ticker").value
        Sheets("Data").Range("a1").CurrentRegion.ClearContents
        
        qurl = "http://ichart.finance.yahoo.com/table.csv?s=" & Symbol
        qurl = qurl & "&a=" & Month(StartDate) - 1 & "&b=" & Day(StartDate) & _
            "&c=" & Year(StartDate) & "&d=" & Month(EndDate) - 1 & "&e=" & _
            Day(EndDate) & "&f=" & Year(EndDate) & "&g=" & Sheets("Data").Range("a1") & "&q=q&y=0&z=" & _
            Symbol & "&x=.csv"
                   
QueryQuote:
             With Sheets("Data").QueryTables.Add(Connection:="URL;" & qurl, Destination:=Sheets("Data").Range("a1"))
                .BackgroundQuery = True
                .TablesOnlyFromHTML = False
                .Refresh BackgroundQuery:=False
                .SaveData = True
            End With
            
            Sheets("Data").Range("a1").CurrentRegion.TextToColumns Destination:=Sheets("Data").Range("a1"), DataType:=xlDelimited, _
                TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
                Semicolon:=False, Comma:=True, Space:=False, other:=False
            
         Sheets("Data").Columns("A:G").ColumnWidth = 12

    LastRow = Sheets("Data").UsedRange.Row - 2 + Sheets("Data").UsedRange.Rows.Count

    Sheets("Data").Sort.SortFields.Add key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Sheets("Data").Sort
        .SetRange Range("A1:G" & LastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
        .SortFields.Clear
    End With

End Sub


Sub CreateCandlestick()
    Dim nRows As Integer
    Dim ch As ChartObject

    nRows = Sheets("Data").UsedRange.Rows.Count

    'Delete existing charts

    For Each ch In Sheets("CandleChart").ChartObjects
        ch.Delete
    Next

    nRows = Sheets("Data").UsedRange.Rows.Count

    'Create candlestick chart

    Dim OHLCChart As ChartObject
    Set OHLCChart = Sheets("CandleChart").ChartObjects.Add(Left:=Range("b8").Left, Width:=400, Top:=Range("b8").Top, Height:=250)

    With OHLCChart.Chart
        .SetSourceData Source:=Sheets("Data").Range("a1:e" & nRows)
        .ChartType = xlStockOHLC
        .HasTitle = True
        .ChartTitle.Text = "Candlestick Chart for " & Sheets("CandleChart").Range("ticker")
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Price"
        .HasLegend = False
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(220, 230, 241)
        .ChartArea.Format.Line.Visible = msoFalse
        .Parent.Name = "OHLC Chart"
    End With

End Sub


