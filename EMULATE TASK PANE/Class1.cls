VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents app As Application
Attribute app.VB_VarHelpID = -1

Private Sub app_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Not UserForm1.cbAutoUpdate Then Exit Sub
    If UserForm1.Visible And Not CheckingCells Then Call UserForm1.UpdateBox
End Sub

Private Sub app_SheetActivate(ByVal Sh As Object)
    If Not UserForm1.cbAutoUpdate Then Exit Sub
    If UserForm1.Visible Then Call UserForm1.UpdateBox
End Sub

