VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, _
        Cancel As Boolean)

    If SaveAsUI = True Then Exit Sub
    Dim wkb As Workbook
    Set wkb = ActiveWorkbook

    Cancel = False

    If MsgBox("Would you like to copy " & vbCrLf _
        & "this worksheet to " & vbCrLf _
        & "a new workbook?", vbQuestion + vbYesNo) = vbYes Then
        Sheets(ActiveSheet.Name).Copy
        wkb.Activate
    End If
End Sub





