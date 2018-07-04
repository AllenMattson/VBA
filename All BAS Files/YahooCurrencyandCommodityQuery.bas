Attribute VB_Name = "YahooCurrencyandCommodityQuery"
Option Explicit

Sub GetIndexForEnergy()
    ' create a web query in the current worksheet
    ' connect to the web, retrieve data, and paste it
    ' in the worksheet as static text

  Sheets.Add
  With ActiveSheet.QueryTables.Add _
    (Connection:="URL;http://finance.yahoo.com/currency-investing", _
          Destination:=Range("A2"))
         '.Name = "IndexForEnergy"
         .BackgroundQuery = True
         '.WebSelectionType = xlSpecifiedTables
         '.WebTables = "*" '"yfi-commodities-energy"
         .WebFormatting = xlWebFormattingAll 'None
         .Refresh BackgroundQuery:=False
         .SaveData = True
  End With
Cells.UnMerge
End Sub




