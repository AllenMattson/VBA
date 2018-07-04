Attribute VB_Name = "ListSelected_PDF_Files"
Sub ListSelected_PDF_Files()
Dim fd As FileDialog
Dim myFile As Variant
Dim lBox As Object

Application.FileDialog(msoFileDialogOpen).Filters.Clear
Set fd = Application.FileDialog(msoFileDialogOpen)
With fd
    .AllowMultiSelect = True
    .Title = "Select the PDF files..."
    .Filters.Clear
    .Filters.Add "PDF files", "*.pdf"
    If .Show Then
        Workbooks.Add
        Set lBox = Worksheets(1).ListBoxes.Add(Left:=20, Top:=60, Height:=40, Width:=300)
        For Each myFile In .SelectedItems
            lBox.AddItem myFile
        Next
        Range("B4").Value = "These are the pdf files you have selected:"
    End If
End With
End Sub
