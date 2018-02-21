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
    Dim cell As Range
    Set Target = Intersect(Target, Target.Parent.UsedRange)
    If Not Target Is Nothing Then
        For Each cell In Target
            cell.Font.Bold = cell.HasFormula
        Next cell
    End If
End Sub


