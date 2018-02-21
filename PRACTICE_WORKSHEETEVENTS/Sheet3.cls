VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As _
    Excel.Range)
  Dim myRange As Range
    
  On Error Resume Next
  Set myRange = Intersect(Range("A1:A10"), Target)
  If Not myRange Is Nothing Then
     MsgBox "Data entry or edits are not permitted."
  End If
End Sub


