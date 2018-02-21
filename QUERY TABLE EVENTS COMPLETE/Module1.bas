Attribute VB_Name = "Module1"
Option Explicit

Public clsQueryEvents As CQueryEvents

Sub Auto_Open()
    Set clsQueryEvents = New CQueryEvents
    Set clsQueryEvents.QTable = Sheet1.QueryTables(1)
End Sub
