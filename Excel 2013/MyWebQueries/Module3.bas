Attribute VB_Name = "Module3"
Option Explicit

Sub GetAccessData()
    Dim strSQL As String
    Dim strConn As String
    
    strSQL = "Select * from Shippers"
    strConn = "ODBC;DSN=MyNorthwind;UID=;PWD=;" & _
        "Database=Northwind 2007"
    Sheets.Add
    With ActiveSheet.QueryTables.Add(Connection:=strConn, _
        Destination:=Range("B1"), Sql:=strSQL)
        .Refresh
    End With
End Sub


Sub GetDelimitedText()

    Dim qtblOutput As QueryTable
    
    Sheets.Add
    Set qtblOutput = ActiveSheet.QueryTables.Add( _
        Connection:="TEXT;C:\Excel2013_HandsOn\NorthEmployees.csv", _
        Destination:=ActiveSheet.Cells(1, 1))
    With qtblOutput
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        '.TextFileOtherDelimiter = "Tab"
        .Refresh
    End With
End Sub
