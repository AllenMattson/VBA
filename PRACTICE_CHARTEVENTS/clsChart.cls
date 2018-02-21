VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents xlChart As Excel.Chart
Attribute xlChart.VB_VarHelpID = -1

Private Sub xlChart_Activate()
    MsgBox "You've activated a chart embedded in  " & _
        ActiveSheet.Name & "."
End Sub


