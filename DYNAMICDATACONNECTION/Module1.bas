Attribute VB_Name = "Module1"
Sub RefreshQuery()

ActiveWorkbook.Connections("Facility Services").OLEDBConnection.Connection = _
"OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source= " & _
ThisWorkbook.Path & _
"\Facility Services.accdb;Mode=ReadWrite"

ActiveWorkbook.Connections("Facility Services").OLEDBConnection.CommandText = _
"SELECT * FROM [Sales_By_Employee] WHERE [Market] = '" & _
Range("C2").Value & "'"

ActiveWorkbook.Connections("Facility Services").Refresh
    
End Sub


Sub ListConnections()
Dim i As Long
Dim Cn As WorkbookConnection

Worksheets.Add
With ActiveSheet.Range("A1:C1")
.Value = Array("Cn Name", "Connection String", "Command Text")
.EntireColumn.AutoFit
End With

For Each Cn In ThisWorkbook.Connections
i = i + 1

Select Case Cn.Type
Case Is = xlConnectionTypeODBC

With ActiveSheet
.Range("A1").Offset(i, 0).Value = Cn.Name
.Range("A1").Offset(i, 1).Value = Cn.ODBCConnection.Connection
.Range("A1").Offset(i, 2).Value = Cn.ODBCConnection.CommandText
End With

Case Is = xlConnectionTypeOLEDB

With ActiveSheet
.Range("A1").Offset(i, 0).Value = Cn.Name
.Range("A1").Offset(i, 1).Value = Cn.OLEDBConnection.Connection
.Range("A1").Offset(i, 2).Value = Cn.OLEDBConnection.CommandText
End With

End Select

Next Cn

End Sub

