Attribute VB_Name = "Database_Queries"
Option Explicit

Sub RunAccessQuery(strQryName As String)
  Dim cat As ADOX.Catalog
  Dim cmd As ADODB.Command
  Dim rst As ADODB.Recordset
  Dim i As Integer
  Dim strPath As String

  strPath = "C:\Excel2013_HandsOn\Northwind.mdb"

  Set cat = New ADOX.Catalog
  cat.ActiveConnection = _
      "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=" & strPath

  Set cmd = cat.Views(strQryName).Command
  Set rst = cmd.Execute

  Sheets(2).Select
  For i = 0 To rst.Fields.Count - 1
      Cells(1, i + 1).Value = rst.Fields(i).Name
  Next
  With ActiveSheet
      .Range("A2").CopyFromRecordset rst
      .Range(Cells(1, 1), _
          Cells(1, rst.Fields.Count)).Font.Bold = True
      .Range("A1").Select
  End With

  Selection.CurrentRegion.Columns.AutoFit
  rst.Close

  Set cmd = Nothing
  Set cat = Nothing
End Sub

Sub RunAccessParamQuery()
  Dim cat As ADOX.Catalog
  Dim cmd As ADODB.Command
  Dim rst As ADODB.Recordset
  Dim i As Integer
  Dim strPath As String
  Dim StartDate As String
  Dim EndDate As String

  strPath = "C:\Excel2013_HandsOn\Northwind.mdb"
  StartDate = "7/1/96"
  EndDate = "7/31/96"

  Set cat = New ADOX.Catalog
  cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source=" & strPath
  Set cmd = cat.Procedures("Employee Sales by Country").Command

  cmd.Parameters("[Beginning Date]") = StartDate
  cmd.Parameters("[Ending Date]") = EndDate

  Set rst = cmd.Execute

  Sheets.Add
  For i = 0 To rst.Fields.Count - 1
      Cells(1, i + 1).Value = rst.Fields(i).Name
  Next
  With ActiveSheet
      .Range("A2").CopyFromRecordset rst
      .Range(Cells(1, 1), Cells(1, rst.Fields.Count)) _
          .Font.Bold = True
      .Range("A1").Select
  End With
  Selection.CurrentRegion.Columns.AutoFit

  rst.Close
  Set cmd = Nothing
  Set cat = Nothing
End Sub


Sub RunAccessFunction()
  Dim objAccess As Object

  On Error Resume Next
  Set objAccess = GetObject(, "Access.Application")

  ' if no instance of Access is open, create a new one
  If objAccess Is Nothing Then
      Set objAccess = CreateObject("Access.Application")
  End If
  MsgBox "For 1000 Spanish pesetas you will get " & _
      objAccess.EuroConvert(1000, "ESP", "EUR") & _
      " euro dollars. "
  Set objAccess = Nothing
End Sub


