VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Excel.Range)
  Application.EnableEvents = False
    Target = UCase(Target)
    Columns(Target.Column).AutoFit
  Application.EnableEvents = True
End Sub


