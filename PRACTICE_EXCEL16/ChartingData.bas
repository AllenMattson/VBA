Attribute VB_Name = "ChartingData"
Option Explicit

Sub ChartData_withADO()
  Dim conn As New ADODB.Connection
  Dim rst As New ADODB.Recordset
  Dim mySheet As Worksheet
  Dim recArray As Variant
  Dim strQueryName As String
  Dim i As Integer
  Dim j As Integer

  strQueryName = "Category Sales for 1997"

  ' Connect with the database
  conn.Open _
      "Provider=Microsoft.Jet.OLEDB.4.0;" _
      & "Data Source=C:\Excel2013_HandsOn\Northwind.mdb;"

  ' Open Recordset based on the SQL statement
      rst.Open "SELECT * FROM [" & strQueryName & "]", conn, _
      adOpenForwardOnly, adLockReadOnly

  Workbooks.Add
  Set mySheet = Worksheets("Sheet1")
  With mySheet.Range("A1")
    recArray = rst.GetRows()
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

  rst.Close
  conn.Close
  Set rst = Nothing
  Set conn = Nothing

  mySheet.Activate
  Charts.Add
  ActiveChart.ChartType = xl3DColumnClustered
  ActiveChart.SetSourceData _
      Source:=mySheet.Cells(1, 1).CurrentRegion, _
      PlotBy:=xlRows
  ActiveChart.Location Where:=xlLocationAsObject, _
      Name:=mySheet.Name

  With ActiveChart
      .HasTitle = True
      .ChartTitle.Characters.Text = strQueryName
      .Axes(xlCategory).HasTitle = True
      .Axes(xlCategory).AxisTitle.Characters.Text = ""
      .Axes(xlValue).HasTitle = True
      .Axes(xlValue).AxisTitle. _
          Characters.Text = mySheet.Range("B1") & "($)"
      .Axes(xlValue).AxisTitle.Orientation = xlUpward
  End With
End Sub



