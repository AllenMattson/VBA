VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Row Selector"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim ColCnt As Long
    Dim rng As Range
    Dim ColWidths As String
    Dim i As Long
    
    ColCnt = ActiveSheet.UsedRange.Columns.Count
    Set rng = ActiveSheet.UsedRange
    With Me.lbxRange
        .ColumnCount = ColCnt
        .RowSource = rng.Offset(1).Resize(rng.Rows.Count - 1).Address
        For i = 1 To .ColumnCount
            ColWidths = ColWidths & rng.Columns(i).Width & ";"
        Next i
        .ColumnWidths = ColWidths
        .ListIndex = 0
    End With
End Sub

Private Sub cmdAll_Click()
    Dim i As Long
    For i = 0 To Me.lbxRange.ListCount - 1
        Me.lbxRange.Selected(i) = True
    Next i
End Sub

Private Sub cmdNone_Click()
    Dim i As Long
    For i = 0 To Me.lbxRange.ListCount - 1
        Me.lbxRange.Selected(i) = False
    Next i
End Sub


Private Sub cmdOK_Click()
    Dim RowRange As Range
    Dim i As Long
    
    For i = 0 To Me.lbxRange.ListCount - 1
        If Me.lbxRange.Selected(i) Then
            If RowRange Is Nothing Then
                Set RowRange = ActiveSheet.UsedRange.Rows(i + 2)
            Else
                Set RowRange = Union(RowRange, ActiveSheet.UsedRange.Rows(i + 2))
            End If
        End If
    Next i
    If Not RowRange Is Nothing Then RowRange.Select
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub lbxRange_Change()
    Me.lblRowCol.Caption = "Row " & Me.lbxRange.ListIndex + 1 + ActiveSheet.UsedRange.Row
End Sub

