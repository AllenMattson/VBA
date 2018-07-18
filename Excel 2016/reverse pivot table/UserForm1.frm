VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Reverse Pivot"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5115
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
    RefEditInput.Text = ActiveCell.CurrentRegion.Address
End Sub

Sub OKButton_Click()
    Dim SummaryTable As Range
    Dim OutputRange As Range

'   Validate ranges
    On Error Resume Next
    Set SummaryTable = Range(RefEditInput.Text)
    If Err.Number <> 0 Then
        MsgBox "Invalid input range.", vbCritical
        Exit Sub
    End If
    
    Set OutputRange = Range(RefEditOutput.Text).Range("A1")
    If Err.Number <> 0 Then
        MsgBox "Invalid output range.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    If SummaryTable.Count = 1 Or SummaryTable.Rows.Count < 3 Then
        MsgBox "Select a cell in the summary table.", vbCritical
        Exit Sub
    End If
    
    Call ReversePivot(SummaryTable, OutputRange, cbCreateTable)
    
    Unload Me
End Sub

Sub ReversePivot(SummaryTable As Range, OutputRange As Range, CreateTable As Boolean)
    Dim r As Long, c As Long
    Dim OutRow As Long, OutCol As Long

'   Convert the range
    OutRow = 2
    Application.ScreenUpdating = False
    OutputRange.Range("A1:C3") = Array("Column1", "Column2", "Column3")
    For r = 2 To SummaryTable.Rows.Count
        For c = 2 To SummaryTable.Columns.Count
            OutputRange.Cells(OutRow, 1) = SummaryTable.Cells(r, 1)
            OutputRange.Cells(OutRow, 2) = SummaryTable.Cells(1, c)
            OutputRange.Cells(OutRow, 3) = SummaryTable.Cells(r, c)
            OutputRange.Cells(OutRow, 3).NumberFormat = SummaryTable.Cells(r, c).NumberFormat
            OutRow = OutRow + 1
        Next c
    Next r

'   Make it a table?
    On Error Resume Next
    If CreateTable Then _
      ActiveSheet.ListObjects.Add xlSrcRange, _
        OutputRange.CurrentRegion, , xlYes
    On Error GoTo 0
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub

