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

Private Sub Workbook_Activate()
    MsgBox "This workbook contains " & _
        ThisWorkbook.Sheets.Count & " sheets."
End Sub

Private Sub Workbook_Deactivate()
    Dim cell As Range
    For Each cell In ActiveSheet.UsedRange
        If Not IsEmpty(cell) Then
            Debug.Print cell.Address & ":" & cell.Value
        End If
    Next
End Sub

Private Sub Workbook_Open()
    ActiveSheet.Range("A1").Value = Format(Now(), _
                            "mm/dd/yyyy")
    Columns("A").AutoFit
End Sub

Private Sub Workbook_BeforeSave(ByVal _
        SaveAsUI As Boolean, Cancel As Boolean)
    If SaveAsUI = True And _
        ThisWorkbook.Path = vbNullString Then
        MsgBox "This document has not yet " _
            & "been saved." & vbCrLf _
        & "The Save As dialog box will be displayed."
    ElseIf SaveAsUI = True Then
        MsgBox "You are not allowed to use " _
        & "the SaveAs option. "
        Cancel = True
    End If
End Sub


Private Sub Workbook_BeforePrint(Cancel As Boolean)
    Dim response As Integer
    response = MsgBox("Do you want to  " & vbCrLf & _
        "print the workbook's full name in the footer?", _
        vbYesNo)
    If response = vbYes Then
        ActiveSheet.PageSetup.LeftFooter = _
            ThisWorkbook.FullName
    Else
        ActiveSheet.PageSetup.LeftFooter = ""
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If MsgBox("Do you want to change " & vbCrLf _
        & " workbook properties before closing?", _
            vbYesNo) = vbYes Then
        Application.Dialogs(xlDialogProperties).Show
    End If
End Sub


