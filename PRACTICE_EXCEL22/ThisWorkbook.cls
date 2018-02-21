VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforePrint(Cancel As Boolean)
  If Worksheets("Sheet1").Range("A1") <> "Monthly Report" Then
    MsgBox "Please enter correct data in cell A1."
    Cancel = True
  End If
End Sub

