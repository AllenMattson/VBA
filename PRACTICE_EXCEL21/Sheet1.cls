VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    Application.CommandBars("Cell").Controls("Sort").Enabled = False
End Sub

Private Sub Worksheet_Deactivate()
    Application.CommandBars("Cell").Controls("Sort").Enabled = True
End Sub


