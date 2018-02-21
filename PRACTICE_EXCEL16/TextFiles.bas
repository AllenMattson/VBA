Attribute VB_Name = "TextFiles"
Option Explicit

Sub CreateTextFile()
  Dim strPath As String
  Dim conn As New ADODB.Connection
  Dim rst As ADODB.Recordset
  Dim strData As String
  Dim strHeader As String
  Dim strSQL As String
  Dim fld As Variant

  strPath = "C:\Excel2013_HandsOn\Northwind.mdb"

  conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
      & "Data Source=" & strPath & ";"

  conn.CursorLocation = adUseClient

  strSQL = "SELECT * FROM Products WHERE UnitPrice > 50"
  Set rst = conn.Execute(CommandText:=strSQL, Options:=adCmdText)

  ' save the recordset as a tab-delimited file
  strData = rst.GetString(StringFormat:=adClipString, _
      ColumnDelimeter:=vbTab, RowDelimeter:=vbCr, _
    nullExpr:=vbNullString)

  For Each fld In rst.Fields
    strHeader = strHeader + fld.Name & vbTab
  Next

  Open "C:\Excel2013_ByExample\ProductsOver50.txt" For Output As #1
  Print #1, strHeader
  Print #1, strData
  Close #1

  rst.Close
  conn.Close

  Set rst = Nothing
  Set conn = Nothing
End Sub

Sub CreateTextFile2()
  Dim conn As New ADODB.Connection
  Dim rst As ADODB.Recordset
  Dim strPath As String
  Dim strData As String
  Dim strHeader As String
  Dim strSQL As String
  Dim fso As Object
  Dim myFile As Object
  Dim fld As Variant

  Set fso = CreateObject("Scripting.FileSystemObject")
  Set myFile = fso.CreateTextFile( _
      "C:\Excel2013_ByExample\ProductsOver20.txt", True)

  strPath = "C:\Excel2013_HandsOn\Northwind 2007.accdb"

  conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;" _
      & "Data Source=" & strPath & ";"
  conn.CursorLocation = adUseClient
  strSQL = "SELECT * FROM Products WHERE [List Price] > 20"

  Set rst = conn.Execute(CommandText:=strSQL, Options:=adCmdText)

  ' save the recordset as a tab-delimited file
  strData = rst.GetString(StringFormat:=adClipString, _
      ColumnDelimeter:=vbTab, RowDelimeter:=vbCr, _
      nullExpr:=vbNullString)

  For Each fld In rst.Fields
      strHeader = strHeader + fld.Name & vbTab
  Next
  With myFile
      .WriteLine strHeader
      .WriteLine strData
      .Close
  End With
End Sub


