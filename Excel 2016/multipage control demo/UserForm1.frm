VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Multipage Controls"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Me.chkFormulaBar.Value = Application.DisplayFormulaBar
    Me.chkGridlines.Value = ActiveWindow.DisplayGridlines
    Me.chkHeaders.Value = ActiveWindow.DisplayHeadings
    Me.chkPageBreaks.Value = ActiveSheet.DisplayPageBreaks
    Me.chkHScrollbar.Value = ActiveWindow.DisplayHorizontalScrollBar
    Me.chkVScrollbar.Value = ActiveWindow.DisplayVerticalScrollBar
    Me.chkSheetTabs.Value = ActiveWindow.DisplayWorkbookTabs
    Me.chkStatusBar.Value = Application.CommandBars("Status Bar").Visible
    Me.scbZoom.Value = ActiveWindow.Zoom
End Sub

Private Sub chkFormulaBar_Click()
    Application.DisplayFormulaBar = Me.chkFormulaBar.Value
End Sub

Private Sub chkGridlines_Click()
    ActiveWindow.DisplayGridlines = Me.chkGridlines.Value
End Sub

Private Sub chkHeaders_Click()
    ActiveWindow.DisplayHeadings = Me.chkHeaders.Value
End Sub

Private Sub chkHScrollbar_Click()
    ActiveWindow.DisplayHorizontalScrollBar = Me.chkHScrollbar.Value
End Sub

Private Sub chkVScrollbar_Click()
    ActiveWindow.DisplayVerticalScrollBar = Me.chkVScrollbar.Value
End Sub

Private Sub chkPageBreaks_Click()
    ActiveSheet.DisplayPageBreaks = Me.chkPageBreaks.Value
End Sub

Private Sub chkSheetTabs_Click()
    ActiveWindow.DisplayWorkbookTabs = Me.chkSheetTabs.Value
End Sub

Private Sub chkStatusBar_Click()
    Application.CommandBars("Status Bar").Visible = Me.chkStatusBar.Value
End Sub

Private Sub scbZoom_Change()
    ActiveWindow.Zoom = Me.scbZoom.Value
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub


