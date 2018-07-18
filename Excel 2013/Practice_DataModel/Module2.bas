Attribute VB_Name = "Module2"
Option Explicit

Sub DataModel_TableChanges()
    Dim strCmdTxt_1 As String
    Dim strCmdTxt_2 As String
    
    strCmdTxt_1 = """Order Details"",""Orders"",""Products"""
    strCmdTxt_2 = strCmdTxt_1 & ",""Customers"",""Employees"""

    With ActiveWorkbook.Connections("Northwind 2007") _
        .OLEDBConnection
        .CommandText = strCmdTxt_2
       '.CommandText = strCmdTxt_1
        .Refresh
    End With
End Sub



