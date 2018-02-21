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

Private Sub cbLargeIcons_Click()
    Call UpdateImages
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Call UpdateImages
End Sub

Sub UpdateImages()
'   Uses CommandButtons rather than Image controls because
'   the image quality is better on a CommandButton
    Dim i As Long, Size As Long
    Dim ImageName As String
    If cbLargeIcons = True Then Size = 32 Else Size = 16
    For i = 1 To 50
        ImageName = ActiveCell.Offset(i - 1, 0).Value
        If ImageName = "" Then
            ActiveSheet.OLEObjects(i).Object.Picture = Nothing
            ActiveSheet.OLEObjects(i).Object.Caption = ""
        Else
            ActiveSheet.OLEObjects(i).Object.Picture = _
              Application.CommandBars.GetImageMso(ImageName, Size, Size)
            ActiveSheet.OLEObjects(i).Object.Caption = ActiveCell.Row + i - 1
        End If
    Next i
End Sub

