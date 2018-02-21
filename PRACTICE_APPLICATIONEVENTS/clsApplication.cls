VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

Private Sub App_NewWorkbook(ByVal _
                                Wb As Workbook)
    Application.DisplayAlerts = False
    If Wb.Sheets.Count = 3 Then
        Sheets(Array(2, 3)).Delete
    End If
    Application.DisplayAlerts = True
End Sub


Private Sub App_SheetSelectionChange(ByVal Sh _
            As Object, ByVal Target As Range)

    If Selection.Count > 1 Or _
            (Selection.Count < 2 And _
            IsEmpty(Target.Value)) Then
        Application.StatusBar = Target.Address
    Else
        Application.StatusBar = Target.Address & _
                "(" & Target.Value & ")"
    End If
End Sub

Private Sub App_WindowActivate(ByVal _
    Wb As Workbook, ByVal Wn As Window)

    Wn.DisplayFormulas = True

End Sub


Private Sub App_WorkbookBeforeClose(ByVal _
            Wb As Workbook, Cancel As Boolean)
    Dim r As Integer
    Dim p As Variant
    
    Sheets.Add
    r = 1
    For Each p In Wb.BuiltinDocumentProperties
      On Error GoTo ErrorHandle
        Cells(r, 1).Value = p.Name & " = " & _
            ActiveWorkbook.BuiltinDocumentProperties _
            .Item(p.Name).Value
        r = r + 1
    Next
    Exit Sub
ErrorHandle:
        Cells(r, 1).Value = p.Name
        Resume Next
End Sub

Private Sub App_WorkbookBeforePrint(ByVal _
            Wb As Workbook, Cancel As Boolean)
    Wb.PrintOut Copies:=2
End Sub

Private Sub App_WorkbookBeforeSave(ByVal _
    Wb As Workbook, ByVal SaveAsUI As Boolean, _
                Cancel As Boolean)
                                               
    If Wb.Path <> vbNullString Then
       ActiveWindow.Caption = Wb.FullName & _
        " [Last Saved: " & Time & "]"
    End If
End Sub

Private Sub App_WorkbookOpen(ByVal _
                                Wb As Workbook)
    If Wb.FileFormat = xlCSV Then
       If MsgBox("Do you want to save this " & vbCrLf _
         & "file as an Excel workbook?", vbYesNo, _
            "Original file format: " _
         & "comma delimited file") = vbYes Then
                Wb.SaveAs FileFormat:=xlWorkbookNormal
        End If
    End If
End Sub



