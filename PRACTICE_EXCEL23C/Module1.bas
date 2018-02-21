Attribute VB_Name = "Module1"
Option Explicit

Sub PivotWithCalcItems()
    Dim strConn As String
    Dim strSQL As String
    Dim myArray As Variant
    Dim destRng As Range
    Dim strPivot As String

    strConn = "Driver={Microsoft Access Driver (*.mdb)};" & _
            "DBQ=" & "C:\Excel2013_ByExample\" & _
            "Northwind.mdb;"

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

    With ActiveSheet.PivotTables(strPivot).PivotFields("CompanyName")
        .Orientation = xlPageField
        .Position = 1
    End With

    With ActiveSheet.PivotTables(strPivot).PivotFields("Country")
        .Orientation = xlRowField
        .Position = 1
    End With

    ActiveSheet.PivotTables(strPivot).AddDataField _
        ActiveSheet.PivotTables(strPivot).PivotFields("ExtendedPrice"), _
        "Sum of ExtendedPrice", xlSum

    With ActiveSheet.PivotTables(strPivot).PivotFields("Salesperson")
        .Orientation = xlRowField
        .Position = 1
    End With

    With ActiveSheet.PivotTables(strPivot).PivotFields("Salesperson")
        .Orientation = xlPageField
        .Position = 1
    End With

    With ActiveSheet.PivotTables(strPivot).PivotFields("Salesperson")
        .Orientation = xlColumnField
        .Position = 1
    End With

    ActiveSheet.PivotTables(strPivot).PivotFields("Country"). _
        CalculatedItems.Add "North America", "=USA+Canada", True
    ActiveSheet.PivotTables(strPivot).PivotFields("Country"). _
        CalculatedItems.Add "South America", _
        "=Argentina+Brazil+Venezuela ", True
    ActiveSheet.PivotTables(strPivot).PivotFields("Country"). _
        CalculatedItems("North America").StandardFormula = _
        "=USA+Canada+Mexico"
    ActiveSheet.PivotTables(strPivot).PivotFields("Country"). _
        CalculatedItems.Add "Europe", _
        "=Austria+Belgium+Denmark+Finland+" & _
        "France+Germany+Ireland+Italy+Norway+Poland+" & _
        "Portugal+Spain+Sweden+Switzerland+UK", True

    With ActiveSheet.PivotTables(strPivot).PivotFields("Country")
        .PivotItems("Argentina").Visible = False
        .PivotItems("Austria").Visible = False
        .PivotItems("Belgium").Visible = False
        .PivotItems("Brazil").Visible = False
        .PivotItems("Canada").Visible = False
        .PivotItems("Denmark").Visible = False
        .PivotItems("Finland").Visible = False
        .PivotItems("France").Visible = False
        .PivotItems("Germany").Visible = False
        .PivotItems("Ireland").Visible = False
        .PivotItems("Italy").Visible = False
        .PivotItems("Mexico").Visible = False
        .PivotItems("Norway").Visible = False
        .PivotItems("Poland").Visible = False
        .PivotItems("Portugal").Visible = False
        .PivotItems("Spain").Visible = False
        .PivotItems("Sweden").Visible = False
        .PivotItems("Switzerland").Visible = False
        .PivotItems("UK").Visible = False
        .PivotItems("USA").Visible = False
        .PivotItems("Venezuela").Visible = False
    End With

    ActiveSheet.PivotTables(strPivot).PivotFields("Country").Caption = _
        "Continent"

'Add this code after running the PivotWithCalcItems procedure
ActiveSheet.PivotTables(strPivot).PivotFields("Salesperson"). _
            CalculatedItems.Add "Male", _
            "=Michael Suyama+Andrew Fuller+Robert King+" & _
            "Steven Buchanan", True

    ActiveSheet.PivotTables(strPivot).PivotFields("Salesperson"). _
        CalculatedItems.Add "Female", _
        "=Anne Dodsworth+Laura Callahan+Janet Leverling+" & _
        "Margaret Peacock+Nancy Davolio", True

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Salesperson")

        .PivotItems("Andrew Fuller").Visible = False
        .PivotItems("Anne Dodsworth").Visible = False
        .PivotItems("Janet Leverling").Visible = False
        .PivotItems("Laura Callahan").Visible = False
        .PivotItems("Margaret Peacock").Visible = False
        .PivotItems("Michael Suyama").Visible = False
        .PivotItems("Nancy Davolio").Visible = False
        .PivotItems("Robert King").Visible = False
        .PivotItems("Steven Buchanan").Visible = False
    End With

    With ActiveSheet.PivotTables(strPivot). _
        PivotFields("Sum of ExtendedPrice").NumberFormat = _
        "$#,##0.00"
    End With

    With ActiveSheet.PivotTables(strPivot).PivotFields("ProductName")
        .Orientation = xlRowField
        .Position = 2
    End With

    ActiveSheet.PivotTables(strPivot). _
        PivotFields("ProductName").Orientation = xlHidden
End Sub


 Sub ListCalcFieldsItems()
    Dim pivTable As PivotTable
    Dim fld As PivotField   ' field enumerator
    Dim itm As PivotItem   ' item enumerator
    Dim r As Integer   ' row number

    Set pivTable = Worksheets(1).PivotTables(1)

    On Error Resume Next

    ' print to the Immediate window the names of fields
    ' and calculated items
    For Each fld In pivTable.PivotFields
      If fld.IsCalculated Then
        Debug.Print fld.Name & ":" & _
        fld.Name & vbTab & "-->Calculated field"
      Else
        Debug.Print fld.Name
      End If
      For Each itm In pivTable. _
        PivotFields(fld.Name).CalculatedItems
          Debug.Print fld.Name & ":" & _
            itm.Name & vbTab & "--> Calculated item"
          ' enter information about Calculated items
          ' in a worksheet
          r = r + 1
          With Worksheets("Sheet2")
            .Cells(r, 1).Value = itm.Name
            .Cells(r, 2).Value = Chr(39) & itm.Formula
          End With
        Next
    Next
End Sub



