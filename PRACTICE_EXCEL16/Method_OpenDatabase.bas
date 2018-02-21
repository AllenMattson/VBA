Attribute VB_Name = "Method_OpenDatabase"
Option Explicit


 Sub OpenAccessDatabase()
      On Error Resume Next

      Workbooks.OpenDatabase _
          Filename:="C:\Excel2013_HandsOn\Northwind.mdb"
  Exit Sub
End Sub

Sub CountCustomersByCountry()
  On Error Resume Next

  Workbooks.OpenDatabase _
      Filename:="C:\Excel2013_HandsOn\Northwind.mdb", _
      CommandText:="Select * from Customers", _
      CommandType:=xlCmdSql, _
      BackgroundQuery:=True, _
      ImportDataAs:=xlPivotTableReport
  Exit Sub
End Sub



