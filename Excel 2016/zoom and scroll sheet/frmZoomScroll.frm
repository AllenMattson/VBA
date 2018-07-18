VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmZoomScroll 
   Caption         =   "Scroll & Zoom Demo"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   OleObjectBlob   =   "frmZoomScroll.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmZoomScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Me.lblZoom.Caption = ActiveWindow.Zoom & "%"
'   Zoom
    With Me.scbZoom
        .Min = 10
        .Max = 400
        .SmallChange = 1
        .LargeChange = 10
        .Value = ActiveWindow.Zoom
    End With
    
'   Horizontally scrolling
    With Me.scbColumns
        .Min = 1
        .Max = ActiveSheet.UsedRange.Columns.Count
        .Value = ActiveWindow.ScrollColumn
        .LargeChange = 25
        .SmallChange = 1
    End With
    
'   Vertically scrolling
    With Me.scbRows
        .Min = 1
        .Max = ActiveSheet.UsedRange.Rows.Count
        .Value = ActiveWindow.ScrollRow
        .LargeChange = 25
        .SmallChange = 1
    End With
End Sub

Private Sub scbZoom_Change()
    With ActiveWindow
        .Zoom = Me.scbZoom.Value
        Me.lblZoom = .Zoom & "%"
        .ScrollColumn = Me.scbColumns.Value
        .ScrollRow = Me.scbRows.Value
    End With
End Sub

Private Sub scbColumns_Change()
    ActiveWindow.ScrollColumn = Me.scbColumns.Value
End Sub

Private Sub scbRows_Change()
    ActiveWindow.ScrollRow = Me.scbRows.Value
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

