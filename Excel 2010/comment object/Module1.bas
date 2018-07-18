Attribute VB_Name = "Module1"
Option Explicit

Sub CountComments()
    Dim CommentCount As Integer
    Dim cell As Range
    Dim x As String
    
'   Exit if no comments
    CommentCount = 0
    For Each cell In ActiveSheet.UsedRange
        On Error Resume Next
        x = cell.Comment.Text
        If Err = 0 Then CommentCount = CommentCount + 1
    Next cell
    If CommentCount = 0 Then
        MsgBox "The active worksheet has no comments.", vbInformation
    Else
        MsgBox "The active worksheet contains " & CommentCount & " comments.", vbInformation
    End If
End Sub

Sub SelectCommentCells()
    Cells.SpecialCells(xlCellTypeComments).Select
End Sub

Sub ToggleComments()
    If Application.DisplayCommentIndicator = xlCommentAndIndicator Then
        Application.DisplayCommentIndicator = xlCommentIndicatorOnly
    Else
        Application.DisplayCommentIndicator = xlCommentAndIndicator
    End If
End Sub
Sub ListComments()
    Dim CommentCount As Integer
    Dim cell As Range
    Dim x As String
    Dim CommentSheet As Worksheet
    Dim OldSheets As Integer
    Dim Row As Integer
    
'   Exit if no comments
    CommentCount = 0
    For Each cell In ActiveSheet.UsedRange
        On Error Resume Next
        x = cell.Comment.Text
        If Err = 0 Then CommentCount = CommentCount + 1
    Next cell
    If CommentCount = 0 Then
        MsgBox "The active worksheet does not contains any comments.", vbInformation
        Exit Sub
    End If
    
'   Create new workbook with one sheet
    On Error GoTo 0
    Set CommentSheet = ActiveSheet
    OldSheets = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = 1
    Workbooks.Add
    Application.SheetsInNewWorkbook = OldSheets
    ActiveWorkbook.Windows(1).Caption = "Comments for " & CommentSheet.Name & " in " & CommentSheet.Parent.Name
    
'   List the comments
    Row = 1
    Cells(Row, 1) = "Address"
    Cells(Row, 2) = "Contents"
    Cells(Row, 3) = "Comment"
    Range(Cells(Row, 1), Cells(Row, 3)).Font.Bold = True
    For Each cell In CommentSheet.UsedRange
        On Error Resume Next
        x = cell.Comment.Text
        If Err = 0 Then
            Row = Row + 1
            Cells(Row, 1) = cell.Address(rowabsolute:=False, columnabsolute:=False)
            Cells(Row, 2) = " " & cell.Formula
            Cells(Row, 3) = cell.Comment.Text
        End If
    Next cell
    Columns("B:B").EntireColumn.AutoFit
    Columns("C:C").ColumnWidth = 34
    Cells.EntireRow.AutoFit
End Sub

Sub ChangeColorofComments()
'   Change colors randomly
    Dim cmt As Comment
    For Each cmt In ActiveSheet.Comments
        cmt.Shape.Fill.ForeColor.SchemeColor = Int((80) * Rnd + 1) '1-80
        cmt.Shape.TextFrame.Characters.Font.ColorIndex = Int((56) * Rnd + 1) '1-56
    Next cmt
End Sub

Sub ChangeShape()
    Dim cmt As Comment
    
    For Each cmt In ActiveSheet.Comments
        If cmt.Shape.AutoShapeType = msoShapeRectangle Then
            cmt.Shape.AutoShapeType = msoShapeRoundedRectangle
        Else
            cmt.Shape.AutoShapeType = msoShapeRectangle
    End If
    Next cmt
End Sub


