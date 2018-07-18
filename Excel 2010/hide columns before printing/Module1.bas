Attribute VB_Name = "Module1"
Sub UnhideColumns()
    Worksheets("Sheet1").Range("B:D").EntireColumn.Hidden = False
End Sub
