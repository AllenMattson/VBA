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
    cbFormulaBar = Application.DisplayFormulaBar
    cbGridlines = ActiveWindow.DisplayGridlines
    cbHeaders = ActiveWindow.DisplayHeadings
    cbPageBreaks = ActiveSheet.DisplayPageBreaks
    cbHScrollbar = ActiveWindow.DisplayHorizontalScrollBar
    cbVScrollbar = ActiveWindow.DisplayVerticalScrollBar
    cbSheetTabs = ActiveWindow.DisplayWorkbookTabs
    cbStatusBar = Application.CommandBars("Status Bar").Visible
    ScrollBarZoom = ActiveWindow.Zoom
End Sub

Private Sub cbFormulaBar_Click()
    Application.DisplayFormulaBar = cbFormulaBar
End Sub

Private Sub cbGridlines_Click()
    ActiveWindow.DisplayGridlines = cbGridlines
End Sub

Private Sub cbHeaders_Click()
    ActiveWindow.DisplayHeadings = cbHeaders
End Sub

Private Sub cbHScrollbar_Click()
    ActiveWindow.DisplayHorizontalScrollBar = cbHScrollbar
End Sub

Private Sub cbVScrollbar_Click()
    ActiveWindow.DisplayVerticalScrollBar = cbVScrollbar
End Sub

Private Sub cbPageBreaks_Click()
    ActiveSheet.DisplayPageBreaks = cbPageBreaks
End Sub

Private Sub cbSheetTabs_Click()
    ActiveWindow.DisplayWorkbookTabs = cbSheetTabs
End Sub

Private Sub cbStatusBar_Click()
    Application.CommandBars("Status Bar").Visible = cbStatusBar
End Sub

Private Sub ScrollBarZoom_Change()
    ActiveWindow.Zoom = ScrollBarZoom
End Sub

Private Sub CloseButton_Click()
    Unload Me
End Sub


