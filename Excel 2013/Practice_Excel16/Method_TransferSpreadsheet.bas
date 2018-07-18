Attribute VB_Name = "Method_TransferSpreadsheet"
Option Explicit

Sub ExportData()
  Dim objAccess As Access.Application
  Set objAccess = CreateObject("Access.Application")

  objAccess.OpenCurrentDatabase filepath:= _
      "C:\Excel2013_HandsOn\Northwind.mdb"

  objAccess.DoCmd.TransferSpreadsheet _
      TransferType:=acExport, _
      SpreadsheetType:=acSpreadsheetTypeExcel12, _
      TableName:="Shippers", _
      Filename:="C:\Excel2013_ByExample\Shippers.xls", _
      HasFieldNames:=True, _
      Range:="Sheet1"

  objAccess.Quit
  Set objAccess = Nothing
End Sub



