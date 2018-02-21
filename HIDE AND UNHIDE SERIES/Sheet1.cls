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
    ActiveSheet.Range("Product_A").EntireColumn.Hidden = _
      Not ActiveSheet.OLEObjects(1).Object.Value
End Sub

Private Sub CheckBox2_Click()
    ActiveSheet.Range("Product_B").EntireColumn.Hidden = _
      Not ActiveSheet.OLEObjects(2).Object.Value
End Sub

Private Sub CheckBox3_Click()
    ActiveSheet.Range("Product_C").EntireColumn.Hidden = _
      Not ActiveSheet.OLEObjects(3).Object.Value
End Sub

Private Sub CheckBox4_Click()
    ActiveSheet.Range("Product_D").EntireColumn.Hidden = _
      Not ActiveSheet.OLEObjects(4).Object.Value
End Sub

Private Sub CheckBox5_Click()
    ActiveSheet.Range("Product_E").EntireColumn.Hidden = _
      Not ActiveSheet.OLEObjects(5).Object.Value
End Sub

Private Sub CheckBox6_Click()
    ActiveSheet.Range("Product_F").EntireColumn.Hidden = _
      Not ActiveSheet.OLEObjects(6).Object.Value
End Sub

Private Sub CheckBox7_Click()
    ActiveSheet.Range("Product_G").EntireColumn.Hidden = _
      Not ActiveSheet.OLEObjects(7).Object.Value
End Sub



