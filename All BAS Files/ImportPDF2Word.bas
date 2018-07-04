Attribute VB_Name = "ImportPDF2Word"
Sub ImportPDF2Word()
  Dim StrFile As String, Rng As Range, DocSrc As Document
  With Application.FileDialog(msoFileDialogOpen)
    .Filters.Clear
    .Filters.Add "PDF Files", "*.pdf"
    .AllowMultiSelect = False
    .Show
    If .SelectedItems.Count = 0 Then Exit Sub
    StrFile = .SelectedItems(1)
  End With
  Set Rng = ActiveDocument.Range.Characters.Last
  Rng.InsertAfter vbCr
  Rng.Collapse wdCollapseEnd
  Set DocSrc = Documents.Open(FileName:=StrFile, AddToRecentFiles:=False)
  With DocSrc
    Rng.FormattedText = .Range.FormattedText
    .Close False
  End With
  Set Rng = Nothing: Set DocSrc = Nothing
End Sub
