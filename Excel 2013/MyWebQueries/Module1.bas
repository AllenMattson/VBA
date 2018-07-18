Attribute VB_Name = "Module1"
Option Explicit

Sub GetIndexForEnergy()
    ' create a web query in the current worksheet
    ' connect to the web, retrieve data, and paste it
    ' in the worksheet as static text

  Sheets.Add
  With ActiveSheet.QueryTables.Add _
    (Connection:="URL;http://finance.yahoo.com/currency-investing", _
          Destination:=Range("A2"))
         .Name = "IndexForEnergy"
         .BackgroundQuery = True
         .WebSelectionType = xlSpecifiedTables
         .WebTables = "yfi-commodities-energy"
         .WebFormatting = xlWebFormattingNone
         .Refresh BackgroundQuery:=False
         .SaveData = True
  End With
End Sub


