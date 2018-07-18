Attribute VB_Name = "QueryTable"
Option Explicit

Sub CreateQueryTable()
  Dim myQryTable As Object
  Dim myDb As String
  Dim strConn As String
  Dim Dest As Range
  Dim strSQL As String

  myDb = "C:\Excel2013_HandsOn\Northwind.mdb"
  strConn = "OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;" _
      & "Data Source=" & myDb & ";"

  Workbooks.Add
  Set Dest = Worksheets(1).Range("A1")
  Sheets(1).Select
  strSQL = "SELECT * FROM Products WHERE UnitPrice > 20"
  Set myQryTable = ActiveSheet.QueryTables.Add(strConn, _
      Dest, _
      strSQL)
  With myQryTable
      .RefreshStyle = xlInsertEntireRows
      .Refresh False
  End With
End Sub



