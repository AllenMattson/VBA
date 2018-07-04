Attribute VB_Name = "ListSelected_PDF_Files"
Private lBox As Object
Public Sub ListSelected_PDF_Files()
Attribute ListSelected_PDF_Files.VB_ProcData.VB_Invoke_Func = "e\n14"
Dim fd As FileDialog
Dim myFile As Variant
'Dim lBox As Object
Dim Shp As Shape
Application.FileDialog(msoFileDialogOpen).Filters.Clear
Set fd = Application.FileDialog(msoFileDialogOpen)
With fd
    .AllowMultiSelect = True
    .Title = "Select the PDF files..."
    .Filters.Clear
    .Filters.Add "PDF files", "*.pdf"
    If .Show Then
        'Workbooks.Add
        Cells.Clear
        For Each Shp In ActiveSheet.Shapes
            Shp.Delete
        Next Shp
        Set lBox = Worksheets(1).ListBoxes.Add(Left:=20, Top:=60, Height:=80, Width:=600)
        For Each myFile In .SelectedItems
            lBox.AddItem myFile
        Next
        With lBox
            '.LinkedCell = "$A$1"
            .Name = "ListBox1"
        End With
        AddButtonScraper
    End If
End With

End Sub
Sub AddButtonScraper()
  Dim btn As Button
  Application.ScreenUpdating = False
  ActiveSheet.Buttons.Delete
  Dim t As Range
    Set t = ActiveSheet.Range("B2:C4")
    Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
    With btn
      .OnAction = "btns"
      .Caption = "Scrape PDF"
      .Name = "Btn1"
    End With
  Application.ScreenUpdating = True
End Sub
Sub btnS()
With Worksheets("Sheet1").ListBoxes("ListBox1")
For i = 1 To .ListCount
    'call macro and pass file name to module to open with word
    If .Selected(i) Then Range("K1") = .List(i) 'MsgBox .List(i)
Next i
End With
End Sub
