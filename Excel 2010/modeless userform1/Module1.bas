Attribute VB_Name = "Module1"
Option Explicit

Sub ShowCellInfoBox()
'   Using two Subs ensures that the file can be opened
'   without a compile error in Excel 97.
    If Val(Application.Version) < 9 Then
        MsgBox "This demo requires Excel 2000 or later.", vbCritical
    Else
        Call ShowModelessForm
    End If
End Sub

Sub ShowModelessForm()
        UserForm1.Show vbModeless
End Sub

Sub UpdateBox()
    With UserForm1
'       Make sure a worksheet is active
        If TypeName(ActiveSheet) <> "Worksheet" Then
            .lblFormula.Caption = "N/A"
            .lblNumFormat.Caption = "N/A"
            .lblLocked.Caption = "N/A"
            Exit Sub
        End If
    
        .Caption = "Cell: " & ActiveCell.Address(False, False)
'       Formula
        If ActiveCell.HasFormula Then
            .lblFormula.Caption = ActiveCell.Formula
        Else
            .lblFormula.Caption = "(none)"
        End If
'       Number format
        .lblNumFormat.Caption = ActiveCell.NumberFormat
'       Locked
        .lblLocked.Caption = ActiveCell.Locked
    End With
End Sub

