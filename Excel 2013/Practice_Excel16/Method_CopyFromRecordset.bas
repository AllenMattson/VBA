Attribute VB_Name = "Method_CopyFromRecordset"
Option Explicit

Sub GetProducts()
    Dim conn As New ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim strPath As String
    Dim j As Integer

    strPath = "C:\Excel2013_HandsOn\Northwind.mdb"

    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0 ;" _
        & "Data Source=" & strPath & ";"
    conn.CursorLocation = adUseClient

    ' Create a Recordset from all the records
    ' in the Products table

    Set rst = conn.Execute(CommandText:="Products", _
        Options:=adCmdTable)

    rst.MoveFirst

    ' transfer the data to Excel
    ' get the names of fields first
    With Worksheets("Sheet3").Range("A1")
        .CurrentRegion.Clear
        For j = 0 To rst.Fields.Count - 1
            .Offset(0, j) = rst.Fields(j).Name
        Next j
        .Offset(1, 0).CopyFromRecordset rst
        .CurrentRegion.Columns.AutoFit
    End With
    rst.Close
    conn.Close

    Set rst = Nothing
    Set conn = Nothing
End Sub



