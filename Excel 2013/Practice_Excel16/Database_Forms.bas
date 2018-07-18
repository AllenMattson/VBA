Attribute VB_Name = "Database_Forms"
Option Explicit

Dim objAccess As Access.Application

Sub DisplayAccessForm()
  Dim strDb As String
  Dim strFrm As String

  strDb = "C:\Excel2013_HandsOn\Northwind 2007.accdb"
  strFrm = "Customer Details"

  Set objAccess = New Access.Application
  With objAccess
      .OpenCurrentDatabase strDb
      .DoCmd.OpenForm strFrm, acNormal
      .DoCmd.Restore
      .Visible = True
  End With
End Sub

