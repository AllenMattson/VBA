VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CheckBox1_Click()
    If CheckBox1 Then
        Call UpdateChart
        ActiveCell.Select
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Excel.Range)
    If CheckBox1 Then Call UpdateChart
End Sub
