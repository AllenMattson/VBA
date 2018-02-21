Attribute VB_Name = "Module1"
Sub Macro97()
'Step 1:  Declare your variables
    Dim MyRange As Excel.Range
    Dim wd As Word.Application
    Dim wdDoc As Word.Document
    Dim WdRange As Word.Range
    
'Step 2:  Copy the defined range
   Sheets("Revenue Table").Range("B4:F10").Copy
    
'Step 3:  Open the target Word document
    Set wd = New Word.Application
    Set wdDoc = wd.Documents.Open _
    (ThisWorkbook.Path & "\" & "PasteTable.docx")
    wd.Visible = True
    
'Step 4:  Set focus on the target bookmark
    Set WdRange = wdDoc.Bookmarks("DataTableHere").Range
    
'Step 5:  Delete the old table and paste new
    On Error Resume Next
    WdRange.Tables(1).Delete
    WdRange.Paste 'paste in the table
    
'Step 6:  Adjust column widths
    WdRange.Tables(1).Columns.SetWidth _
    (MyRange.Width / MyRange.Columns.Count), wdAdjustSameWidth
    
'Step 7:  Reinsert the bookmark
    wdDoc.Bookmarks.Add "DataTableHere", WdRange

'Step 8:  Memory cleanup
    Set wd = Nothing
    Set wdDoc = Nothing
    Set WdRange = Nothing
    
End Sub

