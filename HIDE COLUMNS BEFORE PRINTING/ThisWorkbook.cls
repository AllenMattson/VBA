VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforePrint(Cancel As Boolean)
    'Hide columns B:D on Sheet1 before printing
    Worksheets("Sheet1").Range("B:D").EntireColumn.Hidden = True
    Application.OnTime Now + TimeValue("0:00:05"), "UnhideColumns"
End Sub

