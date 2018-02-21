Attribute VB_Name = "Module1"
Option Explicit

Sub ExportToHTML()
    Dim Filename As Variant
    Dim TDOpenTag As String, TDCloseTag As String
    Dim CellContents As String
    Dim Rng As Range
    Dim r As Long, c As Integer
    
'   Use the selected range of cells
    Set Rng = Application.Intersect(ActiveSheet.UsedRange, Selection)
    If Rng Is Nothing Then
        MsgBox "Nothing to export.", vbCritical
        Exit Sub
    End If
    
'   Get a file name
    Filename = Application.GetSaveAsFilename( _
        InitialFileName:="myrange.htm", _
        fileFilter:="HTML Files(*.htm), *.htm")
    If Filename = False Then Exit Sub
    
'   Open the text file
    Open Filename For Output As #1
    
'   Write the tags
    Print #1, "<HTML>"
    Print #1, "<TABLE BORDER=1 CELLPADDING=3>"
    
'   Loop through the cells
    For r = 1 To Rng.Rows.Count
        Print #1, "<TR>"
        For c = 1 To Rng.Columns.Count
            Select Case Rng.Cells(r, c).HorizontalAlignment
                Case xlHAlignLeft
                    TDOpenTag = "<TD ALIGN=LEFT>"
                Case xlHAlignCenter
                    TDOpenTag = "<TD ALIGN=CENTER>"
                Case xlHAlignGeneral
                    If IsNumeric(Rng.Cells(r, c)) Then
                      TDOpenTag = "<TD ALIGN=RIGHT>"
                    Else
                      TDOpenTag = "<TD ALIGN=LEFT>"
                    End If
                Case xlHAlignRight
                    TDOpenTag = "<TD ALIGN=RIGHT>"
            End Select
            
            TDCloseTag = "</TD>"
            If Rng.Cells(r, c).Font.Bold Then
                TDOpenTag = TDOpenTag & "<B>"
                TDCloseTag = "</B>" & TDCloseTag
            End If
            If Rng.Cells(r, c).Font.Italic Then
                TDOpenTag = TDOpenTag & "<I>"
                TDCloseTag = "</I>" & TDCloseTag
            End If
            CellContents = Rng.Cells(r, c).Text
            Print #1, TDOpenTag & CellContents & TDCloseTag
        Next c
        Print #1, "</TR>"
    Next r
'   Close the table
    Print #1, "</TABLE>"
    Print #1, "</HTML>"

'   Close the file
    Close #1
    
'   Tell the user
    MsgBox Rng.Count & " cells were exported to " & Filename
End Sub

