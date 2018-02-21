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
    Dim ColCnt As Integer
    Dim rng As Range
    Dim cw As String
    Dim c As Integer
    
    ColCnt = ActiveSheet.UsedRange.Columns.Count
    Set rng = ActiveSheet.UsedRange
    With ListBox1
        .ColumnCount = ColCnt
        .RowSource = rng.Address
        cw = ""
        For c = 1 To .ColumnCount
            cw = cw & rng.Columns(c).Width & ";"
        Next c
        .ColumnWidths = cw
        .ListIndex = 0
    End With
End Sub

Private Sub SelectAllButton_Click()
    Dim r As Integer
    For r = 0 To ListBox1.ListCount - 1
        ListBox1.Selected(r) = True
    Next r
End Sub

Private Sub SelectNoneButton_Click()
    Dim r As Integer
    For r = 0 To ListBox1.ListCount - 1
        ListBox1.Selected(r) = False
    Next r
End Sub


Private Sub OKButton_Click()
    Dim RowRange As Range
    Dim RowCnt As Integer
    Dim r As Integer
    
    RowCnt = 0
    For r = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(r) Then
            RowCnt = RowCnt + 1
            If RowCnt = 1 Then
                Set RowRange = ActiveSheet.UsedRange.Rows(r + 1)
            Else
                Set RowRange = Union(RowRange, ActiveSheet.UsedRange.Rows(r + 1))
            End If
        End If
    Next r
    If Not RowRange Is Nothing Then RowRange.Select
    Unload Me
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub ListBox1_Change()
    lblRowCol = "Row " & ListBox1.ListIndex + 1 + ActiveSheet.UsedRange.Row
End Sub

