Attribute VB_Name = "ExcelToAccess"
Option Explicit

Sub LinkExcel_ToAccess()
  Dim objAccess As Access.Application
  Dim strTableName As String
  Dim strBookName As String
  Dim strPath As String
  Dim strName As String
  
  On Error Resume Next
    
  strPath = ActiveWorkbook.Path
  strBookName = strPath & "\Practice_Excel16.xlsm"
  strName = "Linked_ExcelSheet"

  Set objAccess = New Access.Application

  With objAccess
      .OpenCurrentDatabase "C:\Excel2013_HandsOn\Northwind 2007.accdb"
      .DoCmd.TransferSpreadsheet acLink, _
          acSpreadsheetTypeExcel12Xml, _
          strName, strBookName, True, "mySheet!A1:D7"
      .Visible = True
  End With
  Set objAccess = Nothing
End Sub

Sub AccessTbl_From_ExcelData()
  Dim conn As ADODB.Connection
  Dim cat As ADOX.Catalog
  Dim myTbl As ADOX.Table
  Dim rstAccess As ADODB.Recordset
  Dim rowCount As Integer
  Dim i As Integer

  On Error GoTo ErrorHandler

  ' connect to Access using ADO
  Set conn = New ADODB.Connection
  conn.Open "Provider = Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source = C:\Excel2013_HandsOn\Northwind.mdb;"

  ' create an empty Access table
  Set cat = New Catalog
  cat.ActiveConnection = conn
  Set myTbl = New ADOX.Table
  myTbl.Name = "TableFromExcel"
  cat.Tables.Append myTbl

  ' add fields (columns) to the table
  With myTbl.Columns
    .Append "School No", adVarWChar, 7
    .Append "Equipment Type", adVarWChar, 15
    .Append "Serial Number", adVarWChar, 15
    .Append "Manufacturer", adVarWChar, 20
  End With
  Set cat = Nothing

  MsgBox "The table structure was created."

  ' open a recordset based on the newly created
  ' Access table

  Set rstAccess = New ADODB.Recordset
  With rstAccess
    .ActiveConnection = conn
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open myTbl.Name
  End With

  ' now transfer data from Excel spreadsheet range

  With Worksheets("mySheet")
    rowCount = Range("A2:D7").Rows.Count

      For i = 2 To rowCount + 1
          With rstAccess
            .AddNew    ' add a new record to an Access table
            .Fields("School No") = Cells(i, 1).Text
            .Fields("Equipment Type") = Cells(i, 2).Value
            .Fields("Serial Number") = Cells(i, 3).Value
            .Fields("Manufacturer") = Cells(i, 4).Value
            .Update    ' update the table record
          End With
      Next i
  End With

  MsgBox "Data from an Excel spreadsheet was loaded into the table."

  ' close the Recordset and Connection object and remove them
  ' from memory
  rstAccess.Close
  conn.Close
  Set rstAccess = Nothing
  Set conn = Nothing

  MsgBox "Open the Northwind database to view the table."
AccessTbl_From_ExcelDataExit:
  Exit Sub
ErrorHandler:
  MsgBox Err.Number & ": " & Err.Description
  Resume AccessTbl_From_ExcelDataExit
End Sub


