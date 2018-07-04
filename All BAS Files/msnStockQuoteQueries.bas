Attribute VB_Name = "msnStockQuoteQueries"
Option Explicit

Sub Portfolio()
    Dim sht As Worksheet
    Dim qryTbl As QueryTable

    ' insert a new worksheet in the current workbook
    Set sht = ThisWorkbook.Worksheets.Add
    ' create a new web query in a worksheet
    Set qryTbl = sht.QueryTables.Add _
    (Connection:="URL;http://moneycentral." & _
    "msn.com/investor/external/excel/quotes.asp?" & _
    "SYMBOL=GOOG&SYMBOL=YHOO", Destination:=sht.Range("A1"))
    ' retrieve data from web page and specify formatting
    ' paste data in a worksheet
    With qryTbl
        .BackgroundQuery = True
        .WebSelectionType = xlSpecifiedTables
        .WebTables = "1"
        .WebFormatting = xlWebFormattingAll 'xlWebFormattingNone
        .Refresh BackgroundQuery:=False
        .SaveData = True
    End With
    ' delete unwanted rows/columns
    With sht
        .Rows("2").Delete
        .Columns("B:C").Delete
        .Rows("5:16").Delete
    End With
End Sub

Sub Portfolio2()
    Dim sht As Worksheet
    Dim qryTbl As QueryTable

    ' insert a new worksheet in the current workbook
    Set sht = ThisWorkbook.Worksheets.Add
    ' create a new web query in a worksheet
    Set qryTbl = sht.QueryTables.Add _
    (Connection:="URL;http://moneycentral." & _
      "msn.com/investor/external/excel/quotes.asp?" & _
      "SYMBOL=[""Enter " & "symbols separated by spaces""]", _
       Destination:=sht.Range("A1"))
    ' retrieve data from web page and specify formatting
    ' paste data in a worksheet
    With qryTbl
        .BackgroundQuery = True
        .WebSelectionType = xlSpecifiedTables
        .WebTables = "1"
        .WebFormatting = xlWebFormattingAll
        .Refresh BackgroundQuery:=False
        .SaveData = True
    End With
    ' delete unwanted rows/columns
    With sht
        .Rows("2").Delete
        .Rows("6:18").Delete
        .Columns("B:C").Delete
    End With
End Sub


