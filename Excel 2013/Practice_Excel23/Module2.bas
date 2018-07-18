Attribute VB_Name = "Module2"
Option Explicit

Sub PivotTable_External1()
    Dim strConn As String
    Dim strQuery_1 As String
    Dim strQuery_2 As String
    Dim myArray As Variant
    Dim destRange As Range
    Dim strPivot As String

    strConn = "Driver={Microsoft Access Driver (*.mdb)};" & _
        "DBQ=" & "C:\Excel2013_ByExample\Northwind.mdb;"

    strQuery_1 = "SELECT Customers.CustomerID, " & _
        "Customers.CompanyName," & _
        "Orders.OrderDate, Products.ProductName, Sum([Order " & _
        "Details].[UnitPrice]*[Quantity]*(1-[Discount])) " & _
            "AS Total " & _
        "FROM Products INNER JOIN ((Customers INNER JOIN Orders " & _
        "ON Customers.CustomerID = "

    strQuery_2 = "Orders.CustomerID) INNER JOIN [Order Details] " & _
        "ON Orders.OrderID = [Order Details].OrderID) ON " & _
        "Products.ProductID = [Order Details].ProductID " & _
        "GROUP BY Customers.CustomerID, Customers.CompanyName, " & _
        "Orders.OrderDate, Products.ProductName;"

    myArray = Array(strConn, strQuery_1, strQuery_2)
    Worksheets.Add

    Set destRange = ActiveSheet.Range("B5")
    strPivot = "PivotFromAccess"

    ActiveSheet.PivotTableWizard _
     SourceType:=xlExternal, _
     SourceData:=myArray, _
     TableDestination:=destRange, _
     TableName:=strPivot, _
     SaveData:=False, _
     BackgroundQuery:=False

    ' Close the PivotTable Field List that appears automatically
    ActiveWorkbook.ShowPivotTableFieldList = False

    ' Add fields to the PivotTable
    With ActiveSheet.PivotTables(strPivot)
    .PivotFields("ProductName").Orientation = xlRowField
    .PivotFields("CompanyName").Orientation = xlRowField
    With .PivotFields("Total")
        .Orientation = xlDataField
        .Function = xlSum
        .NumberFormat = "$#,##0.00"
    End With
    .PivotFields("CustomerID").Orientation = xlPageField
    .PivotFields("OrderDate").Orientation = xlPageField
    End With
    ' Autofit columns so all headings are visible
    ActiveSheet.UsedRange.Columns.AutoFit
End Sub



