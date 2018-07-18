Attribute VB_Name = "Examples_ADO"
Option Explicit

Sub ADO_OpenDatabase(strDbPathName)
  Dim con As New ADODB.Connection
  Dim rst As New ADODB.Recordset
  Dim fld As ADODB.Field
  Dim iCol As Integer
  Dim wks As Worksheet
    
  ' connect with the database
  If Right(strDbPathName, 3) = "mdb" Then
      con.Open _
      "Provider=Microsoft.Jet.OLEDB.4.0;" _
          & "Data Source=" & strDbPathName
  ElseIf Right(strDbPathName, 5) = "accdb" Then
      con.Open _
      "Provider = Microsoft.ACE.OLEDB.12.0;" _
       & "Data Source=" & strDbPathName
  Else
      MsgBox "Incorrect filename extension"
      Exit Sub
  End If

  ' open Recordset based on the SQL statement
    rst.Open "SELECT * FROM Employees " & _
      "WHERE City = 'Redmond'", con, _
      adOpenForwardOnly, adLockReadOnly
    
  ' enter data into an Excel worksheet in a new workbook
  Workbooks.Add
  Set wks = ActiveWorkbook.Sheets(1)
  wks.Activate
    
  'write column names to the first worksheet row
  For iCol = 0 To rst.Fields.Count - 1
      wks.Cells(1, iCol + 1).Value = rst.Fields(iCol).Name
  Next
      
  'copy records to the worksheet
  wks.Range("A2").CopyFromRecordset rst
    
  'autofit the columns to make the data fit
  wks.Columns.AutoFit
    
  'release object variables
  Set wks = Nothing
      
  ' close the Recordset and connection with Access
  rst.Close
  con.Close

  ' destroy object variables to reclaim the resources
  Set rst = Nothing
  Set con = Nothing
End Sub


Sub CreateDB_ViaADO()
  Dim cat As ADOX.Catalog
  Set cat = New ADOX.Catalog

  cat.Create "Provider=Microsoft.ACE.OLEDB.12.0;" & _
      "Data Source=C:\Excel2013_ByExample\ExcelDump2.accdb;"

  Set cat = Nothing
End Sub


