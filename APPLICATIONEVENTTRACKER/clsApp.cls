VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents XL As Excel.Application
Attribute XL.VB_VarHelpID = -1

Sub LogEvent(txt)
    EventNum = EventNum + 1
    With UserForm1
        With .lblEvents
            .AutoSize = False
            .Caption = .Caption & vbCrLf & txt
            .Width = UserForm1.FrameEvents.Width - 20
            .AutoSize = True
        End With
        .FrameEvents.ScrollHeight = .lblEvents.Height + 20
        .FrameEvents.ScrollTop = EventNum * 20
    End With
End Sub

Private Sub XL_NewWorkbook(ByVal Wb As Excel.Workbook)
    LogEvent "NewWorkbook: " & Wb.Name
End Sub

Private Sub XL_SheetActivate(ByVal Sh As Object)
    LogEvent "SheetActivate: " & Sh.Name
End Sub

Private Sub XL_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Excel.Range, Cancel As Boolean)
    LogEvent "SheetBeforeDoubleClick: " & Target.Address(False, False) & " in " & Sh.Name
End Sub

Private Sub XL_SheetBeforeRightClick(ByVal Sh As Object, ByVal Target As Excel.Range, Cancel As Boolean)
    LogEvent "SheetBeforeRightClick: " & Target.Address(False, False) & " in " & Sh.Name
End Sub

Private Sub XL_SheetCalculate(ByVal Sh As Object)
    LogEvent "SheetCalculate: " & Sh.Name
End Sub

Private Sub XL_SheetChange(ByVal Sh As Object, ByVal Target As Excel.Range)
    LogEvent "SheetChange: " & Target.Address(False, False) & " in " & Sh.Name
End Sub

Private Sub XL_SheetDeactivate(ByVal Sh As Object)
    LogEvent "SheetDeactivate: " & Sh.Name
End Sub

Private Sub XL_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    LogEvent "SheetFollowHyperlink: " & Target.Name & " in " & Sh.Name
End Sub

Private Sub XL_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Excel.Range)
    LogEvent "SheetSelectionChange: " & Target.Address(False, False) & " in " & Sh.Name
End Sub

Private Sub XL_WindowActivate(ByVal Wb As Excel.Workbook, ByVal Wn As Excel.Window)
    LogEvent "WindowActivate: " & Wn.Caption & " in " & Wb.Name
End Sub

Private Sub XL_WindowDeactivate(ByVal Wb As Excel.Workbook, ByVal Wn As Excel.Window)
    LogEvent "WindowDeactivate: " & Wn.Caption & " in " & Wb.Name
End Sub

Private Sub XL_WindowResize(ByVal Wb As Excel.Workbook, ByVal Wn As Excel.Window)
    LogEvent "WindowResize: " & Wn.Caption & " in " & Wb.Name
End Sub

Private Sub XL_WorkbookActivate(ByVal Wb As Excel.Workbook)
    LogEvent "WorkbookActivate: " & Wb.Name
End Sub

Private Sub XL_WorkbookAddinInstall(ByVal Wb As Excel.Workbook)
    LogEvent "WorkbookAddinInstall: " & Wb.Name
End Sub

Private Sub XL_WorkbookAddinUninstall(ByVal Wb As Excel.Workbook)
    LogEvent "WorkbookAddinUninstall: " & Wb.Name
End Sub

Private Sub XL_WorkbookBeforeClose(ByVal Wb As Excel.Workbook, Cancel As Boolean)
    LogEvent "WorkbookBeforeClose: " & Wb.Name
End Sub

Private Sub XL_WorkbookBeforePrint(ByVal Wb As Excel.Workbook, Cancel As Boolean)
    LogEvent "WorkbookBeforePrint: " & Wb.Name
End Sub

Private Sub XL_WorkbookBeforeSave(ByVal Wb As Excel.Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    LogEvent "WorkbookBeforeSave: " & Wb.Name
End Sub

Private Sub XL_WorkbookDeactivate(ByVal Wb As Excel.Workbook)
    LogEvent "WorkbookDeactivate: " & Wb.Name
End Sub

Private Sub XL_WorkbookNewSheet(ByVal Wb As Excel.Workbook, ByVal Sh As Object)
    LogEvent "WorkbookNewSheet: " & Sh.Name & " in " & Wb.Name
End Sub

Private Sub XL_WorkbookOpen(ByVal Wb As Excel.Workbook)
    LogEvent "WorkbookOpen: " & Wb.Name
End Sub

