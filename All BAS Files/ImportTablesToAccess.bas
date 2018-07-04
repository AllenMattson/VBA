Attribute VB_Name = "ImportTablesToAccess"
Option Compare Database

Private Sub Table()
Dim LPosition As Integer
Dim strFileName As String
Dim myFileName As String
Dim arr() As String
Dim dlg As FileDialog
Set dlg = Application.FileDialog(msoFileDialogFilePicker)

With dlg
.Title = "Select the Excel file to import"
.AllowMultiSelect = True
.Filters.Clear
.Filters.Add "Excel Files", "*.xls*", 1
.Filters.Add "All Files", "*.*", 2
.Show
    'Display paths of each file selected
    For Each File In .SelectedItems
    
    strFileName = .SelectedItems(1)
    myFileName = File
    myFileName = Right(myFileName, Len(myFileName) - InStrRev(myFileName, "Zip_") + 1) 'Name of Table is Document Name
    myFileName = Left(myFileName, Len(myFileName) - 5)
    Debug.Print myFileName
        MsgBox arr(i)
        myFileName = LBound(arr)

    
        myFileName = myFileName - Right(myFileName, LPosition)
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel8, myFileName, strFileName, True
    Next


End With
End Sub
