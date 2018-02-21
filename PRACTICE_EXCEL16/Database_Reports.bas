Attribute VB_Name = "Database_Reports"
Option Explicit


Dim objAccess As Access.Application

Sub DisplayAccessReport()
  Dim strDb As String
  Dim strRpt As String
  strDb = "C:\Excel2013_HandsOn\Northwind.mdb"
  strRpt = "Products by Category"

  Set objAccess = New Access.Application
  With objAccess
      .OpenCurrentDatabase (strDb)
      .DoCmd.OpenReport strRpt, acViewPreview
      .DoCmd.Maximize
      .Visible = True
  End With
End Sub

Sub DisplayAccessReport2(strDb As String, _
    strRpt As String)

  Set objAccess = New Access.Application

  With objAccess
      .OpenCurrentDatabase (strDb)
      .DoCmd.OpenReport strRpt, acViewPreview
      .DoCmd.Maximize
      .Visible = True
  End With
End Sub


' Enter the following procedure in the Code window and run it

Sub ShowReport()
  Dim strDb As String
  Dim strRpt As String

  strDb = InputBox("Enter the name of the database (full path): ")
  strRpt = InputBox("Enter the name of the report:")
  Call DisplayAccessReport2(strDb, strRpt)
End Sub

