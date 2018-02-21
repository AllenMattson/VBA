Attribute VB_Name = "Module1"
Sub Macro98()
'Step 1:  Declare your variables
    Dim wd As Word.Application
    Dim wdDoc As Word.Document
    Dim MyRange As Excel.Range
    Dim MyCell As Excel.Range
    Dim txtAddress As String
    Dim txtCity As String
    Dim txtState As String
    Dim txtPostalCode As String
    Dim txtFname As String
    Dim txtFullname As String
    
    
'Step 2:  Start Word and add a new document
    Set wd = New Word.Application
    Set wdDoc = wd.Documents.Add
    wd.Visible = True
    
    
'Step 3:  Set the range of your contact list
    Set MyRange = Sheets("Contact List").Range("A5:A24")
    
    
'Step 4:  Start the loop through each cell
    For Each MyCell In MyRange.Cells
    
    
'Step 5:  Assign values to each component of the letter
    txtAddress = MyCell.Value
    txtCity = MyCell.Offset(, 1).Value
    txtState = MyCell.Offset(, 2).Value
    txtPostalCode = MyCell.Offset(, 3).Value
    txtFname = MyCell.Offset(, 5).Value
    txtFullname = MyCell.Offset(, 6).Value
    
    
'Step 6:Insert the structure of template document
    wd.Selection.InsertFile _
    ThisWorkbook.Path & "\" & "MailMerge.docx"
    
    
'Step 7:  Fill each relevant bookmark with respective value
    wd.Selection.Goto What:=wdGoToBookmark, Name:="Customer"
    wd.Selection.TypeText Text:=txtFullname
    
    wd.Selection.Goto What:=wdGoToBookmark, Name:="Address"
    wd.Selection.TypeText Text:=txtAddress
    
    wd.Selection.Goto What:=wdGoToBookmark, Name:="City"
    wd.Selection.TypeText Text:=txtCity
    
    wd.Selection.Goto What:=wdGoToBookmark, Name:="State"
    wd.Selection.TypeText Text:=txtState
    
    wd.Selection.Goto What:=wdGoToBookmark, Name:="Zip"
    wd.Selection.TypeText Text:=txtPostalCode
    
    wd.Selection.Goto What:=wdGoToBookmark, Name:="FirstName"
    wd.Selection.TypeText Text:=txtFname
    
    
'Step 8:  Clear any remaining bookmarks
    On Error Resume Next
    wdDoc.Bookmarks("Address").Delete
    wdDoc.Bookmarks("Customer").Delete
    wdDoc.Bookmarks("City").Delete
    wdDoc.Bookmarks("State").Delete
    wdDoc.Bookmarks("FirstName").Delete
    wdDoc.Bookmarks("Zip").Delete
    
    
'Step 9:  Go to the end, insert new page, and start with the next cell
    wd.Selection.EndKey Unit:=wdStory
    wd.Selection.InsertBreak Type:=wdPageBreak
    Next MyCell
    
        
'Step 10:  Set cursor to beginning and clean up memory
    wd.Selection.HomeKey Unit:=wdStory
    wd.Activate
    Set wd = Nothing
    Set wdDoc = Nothing
    
End Sub

