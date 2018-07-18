VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents CardPicture As MSForms.Image
Attribute CardPicture.VB_VarHelpID = -1

Private Sub CardPicture_Click()
    Call GameForm.Card_Click(CardPicture.Name)
End Sub

