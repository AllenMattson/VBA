Attribute VB_Name = "Module3"
Option Explicit

Sub Pivot_External2()
    Dim objPivotCache As PivotCache
    Dim conn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim dbPath As String
    Dim strSQL As String

    dbPath = "C:\Excel2013_HandsOn\Northwind 2007.accdb"

    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" _
            & "Data Source=" & dbPath & _
            "; Persist Security Info=False;"

    strSQL = "SELECT Products.[Product Name], " & _
             "Orders.[Order Date], " & _
             "Sum([Unit Price]*[Quantity]) AS Amount " & _
             "FROM Orders INNER JOIN (Products INNER JOIN " & _
             "[Order Details] ON Products.ID = " & _
             "[Order Details].[Product ID]) ON " & _
             "Orders.[Order ID] = [Order Details].[Order ID] " & _
            "GROUP BY Products.[Product Name], " & _
             "Orders.[Order Date], Products.[Product Name]" & _
             "ORDER BY Sum([Unit Price]*[Quantity]) DESC , " & _
             "Products.[Product Name];"

    Set rst = conn.Execute(strSQL)

    ' Create a PivotTable cache and report
    Set objPivotCache = ActiveWorkbook.PivotCaches.Add( _
        SourceType:=xlExternal)
    Set objPivotCache.Recordset = rst

    Worksheets.Add
    With objPivotCache
        .CreatePivotTable TableDestination:=Range("B6"), _
            TableName:="Invoices"
    End With

    ' Add fields to the PivotTable
    With ActiveSheet.PivotTables("Invoices")
        .SmallGrid = False
        With .PivotFields("Product Name")
             .Orientation = xlRowField
             .Position = 1
        End With
        With .PivotFields("Order Date")
        .Orientation = xlRowField
            .Position = 2
            .Name = "Date"
        End With
        With .PivotFields("Amount")
             .Orientation = xlDataField
             .Position = 1
             .NumberFormat = "$#,##0.00"
        End With
    End With

    ' Autofit columns so all headings are visible
    ActiveSheet.UsedRange.Columns.AutoFit

    ' Clean up
    rst.Close
    conn.Close
    Set rst = Nothing
    Set conn = Nothing

    ' Obtain information about PivotCache
    With ActiveSheet.PivotTables("Invoices").PivotCache
        Debug.Print "Information about the PivotCache"
        Debug.Print "Number of Records: " & .RecordCount
        Debug.Print "Data was last refreshed on: " & .RefreshDate
        Debug.Print "Data was last refreshed by: " & .RefreshName
        Debug.Print "Memory used by PivotCache: " & .MemoryUsed & _
            " (bytes)"
    End With
    
    ActiveSheet.PivotTables("Invoices").PivotCache.RefreshOnFileOpen = True


End Sub



