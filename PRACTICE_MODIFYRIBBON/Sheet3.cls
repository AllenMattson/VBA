VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Activate()
    blnEnabled = False
    objRibbon.Invalidate
    MsgBox "Formula bar checkmark is disabled in this sheet only."
End Sub

Private Sub Worksheet_Deactivate()
    blnEnabled = True
    objRibbon.Invalidate
End Sub

