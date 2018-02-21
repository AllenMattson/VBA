Attribute VB_Name = "Method_GetRows"
Option Explicit

Sub GetData_withGetRows()
  Dim db As DAO.Database
  Dim qdf As DAO.QueryDef
  Dim rst As DAO.Recordset
  Dim recArray As Variant
  Dim i As Integer
  Dim j As Integer
  Dim strPath As String
  Dim a As Variant
  Dim countR As Long
  Dim strShtName As String

  strPath = "C:\Excel2013_HandsOn\Northwind.mdb"
  strShtName = "Returned records"

  Set db = OpenDatabase(strPath)
  Set qdf = db.QueryDefs("Invoices")
  Set rst = qdf.OpenRecordset

  rst.MoveLast
  countR = rst.RecordCount
  a = InputBox("This recordset contains " & _
      countR & " records." & vbCrLf _
      & "Enter number of records to return: ", _
      "Get Number of Records")

  If a = "" Or a = 0 Then Exit Sub
  If a > countR Then
      a = countR
      MsgBox "The number you entered is too large." & vbCrLf _
          & "All records will be returned."
  End If

  Workbooks.Add
  ActiveWorkbook.Worksheets(1).Name = strShtName
  rst.MoveFirst
      With Worksheets(strShtName).Range("A1")
        .CurrentRegion.Clear
        recArray = rst.GetRows(a)
        For i = 0 To UBound(recArray, 2)
            For j = 0 To UBound(recArray, 1)
                .Offset(i + 1, j) = recArray(j, i)
            Next j
        Next i
        For j = 0 To rst.Fields.Count - 1
            .Offset(0, j) = rst.Fields(j).Name
            .Offset(0, j).EntireColumn.AutoFit
        Next j
      End With
  db.Close
End Sub


Sub GetData_withGetRows_ADO()
  Dim cat As ADOX.Catalog
  Dim cmd As ADODB.Command
  Dim rst As ADODB.Recordset
  Dim strConnect As String
  Dim recArray As Variant
  Dim i As Integer
  Dim j As Integer
  Dim strPath As String
  Dim a As Variant
  Dim countR As Long
  Dim strShtName As String

  strConnect = "Provider=Microsoft.ACE.OLEDB.12.0;" _
      & "Data Source=C:\Excel2013_HandsOn\Northwind 2007.accdb;"

  strShtName = "Returned records"

  Set cat = New ADOX.Catalog
  cat.ActiveConnection = strConnect

  Set cmd = cat.Views("Order Summary").Command
  Set rst = New ADODB.Recordset
  rst.Open cmd, , adOpenStatic, adLockReadOnly

  countR = rst.RecordCount
  a = InputBox("This recordset contains " & _
      countR & " records." & vbCrLf _
      & "Enter number of records to return: ", _
      "Get Number of Records")

  If a = "" Or a = 0 Then Exit Sub
  If a > countR Then
      a = countR
      MsgBox "The number you entered is too large." & vbCrLf _
          & "All records will be returned."
  End If

  Workbooks.Add
  ActiveWorkbook.Worksheets(1).Name = strShtName
  rst.MoveFirst
      With Worksheets(strShtName).Range("A1")
        .CurrentRegion.Clear
        recArray = rst.GetRows(a)
        For i = 0 To UBound(recArray, 2)
            For j = 0 To UBound(recArray, 1)
                .Offset(i + 1, j) = recArray(j, i)
            Next j
        Next i
        For j = 0 To rst.Fields.Count - 1
            .Offset(0, j) = rst.Fields(j).Name
            .Offset(0, j).EntireColumn.AutoFit
        Next j
      End With

  Set rst = Nothing
  Set cmd = Nothing
  Set cat = Nothing
End Sub

