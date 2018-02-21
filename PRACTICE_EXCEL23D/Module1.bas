Attribute VB_Name = "Module1"
Option Explicit

Sub GeneratePivotReport()
    Dim strConn As String
    Dim strSQL As String
    Dim myArray As Variant
    Dim destRng As Range
    Dim strPivot As String

    strConn = "Driver={Microsoft Access Driver (*.mdb)};" & _
        "DBQ=" & "C:\Excel2013_ByExample\Northwind.mdb;"

    strSQL = "SELECT Invoices.Customers.CompanyName, " & _
        "Invoices.Country, Invoices.Salesperson, " & _
        "Invoices.ProductName, Invoices.ExtendedPrice " & _
        "FROM Invoices ORDER BY Invoices.Country"

    myArray = Array(strConn, strSQL)
    Worksheets.Add

    Set destRng = ActiveSheet.Range("B5")
    strPivot = "PivotTable1"

    ActiveSheet.PivotTableWizard _
        SourceType:=xlExternal, _
        SourceData:=myArray, _
        TableDestination:=destRng, _
        TableName:=strPivot, _
        SaveData:=False, _
        BackgroundQuery:=False

    With ActiveSheet.PivotTables(strPivot).PivotFields("ProductName")
        .Orientation = xlPageField
        .Position = 1
    End With


    With ActiveSheet.PivotTables(strPivot).PivotFields("Country")
        .Orientation = xlRowField
        .Position = 1
    End With

    With ActiveSheet.PivotTables(strPivot).PivotFields("Salesperson")
        .Orientation = xlColumnField
        .Position = 1
    End With

    ActiveSheet.PivotTables(strPivot).AddDataField _
    ActiveSheet.PivotTables(strPivot).PivotFields("ExtendedPrice"), _
    "Sum of ExtendedPrice", xlSum

    With ActiveSheet.PivotTables(strPivot). _
    PivotFields("Sum of ExtendedPrice").NumberFormat = "$#,##0.00"
    End With
End Sub

Sub CreatePivotChart()
    Dim shp As Shape
    Dim rngSource As Range
    Dim pvtTable As PivotTable
    Dim r As Integer

    Set pvtTable = Worksheets("Sheet2").PivotTables(1)

    ' set the current page for the PivotTable report to the
    ' page named "Tofu"
    pvtTable.PivotFields("ProductName").CurrentPage = "Tofu"

    Set rngSource = pvtTable.TableRange2
    Set shp = ActiveSheet.Shapes.AddChart

    shp.Chart.SetSourceData Source:=rngSource
    shp.Chart.SetElement (msoElementChartTitleAboveChart)
    shp.Chart.ChartTitle.Caption = _
        pvtTable.PivotFields("ProductName").CurrentPage

    r = ActiveSheet.UsedRange.Rows.Count + 3

    With Range("B" & r & ":E" & r + 15)
        shp.Width = .Width
        shp.Height = .Height
        shp.Left = .Left
        shp.Top = .Top
    End With
End Sub

