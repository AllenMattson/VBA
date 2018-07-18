Attribute VB_Name = "DialogBoxes"
Option Explicit

Sub ListFilters()
  Dim fdf As FileDialogFilters
  Dim fltr As FileDialogFilter
  Dim c As Integer

  Set fdf = Application.FileDialog(msoFileDialogOpen).Filters

  Workbooks.Add
  Cells(1, 1).Select
  Selection.Formula = "List of Default filters"
  With fdf
    c = .Count
    For Each fltr In fdf
      Selection.Offset(1, 0).Formula = fltr.Description & _
          ": " & fltr.Extensions
      Selection.Offset(1, 0).Select
    Next
  MsgBox c & " filters were written to a worksheet."
  End With
End Sub


Sub ListFilters2()
  Dim fdf As FileDialogFilters
  Dim fltr As FileDialogFilter
  Dim c As Integer

  Set fdf = Application.FileDialog(msoFileDialogOpen).Filters

  Workbooks.Add
  Cells(1, 1).Select
  Selection.Formula = "List of Default filters"
  With fdf
    c = .Count
    For Each fltr In fdf
      Selection.Offset(1, 0).Formula = fltr.Description & _
          ": " & fltr.Extensions
      Selection.Offset(1, 0).Select
    Next
    MsgBox c & " filters were written to a worksheet."
    .Add "Temporary Files", "*.tmp", 1
    c = .Count
    MsgBox "There are now " & c & " filters." & vbCrLf _
        & "Check for yourself."
    Application.FileDialog(msoFileDialogOpen).Show
  End With
End Sub


Sub ListSelectedFiles()
  Dim fd As FileDialog
  Dim myFile As Variant
  Dim lbox As Object

  Application.FileDialog(msoFileDialogOpen).Filters.Clear
  Set fd = Application.FileDialog(msoFileDialogOpen)
  With fd
      .AllowMultiSelect = True
      If .Show Then
        Workbooks.Add
      Set lbox = Worksheets(1).Shapes. _
          AddFormControl(xlListBox, _
          Left:=20, Top:=60, Height:=40, Width:=300)
      lbox.ControlFormat.MultiSelect = xlNone
      For Each myFile In .SelectedItems
        lbox.ControlFormat.AddItem myFile
      Next
      Range("B4").Formula = _
          "You've selected the following " & _
          lbox.ControlFormat.ListCount & " files:"
      lbox.ControlFormat.ListIndex = 1
    End If
  End With
End Sub

Sub OpenRightAway()
  Dim fd As FileDialog
  Dim myFile As Variant

  Set fd = Application.FileDialog(msoFileDialogOpen)
  With fd
    .AllowMultiSelect = True
    If .Show Then
      For Each myFile In .SelectedItems
        .Execute
      Next
    End If
  End With
End Sub


