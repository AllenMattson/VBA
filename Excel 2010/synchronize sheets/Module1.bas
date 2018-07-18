Attribute VB_Name = "Module1"
Sub SynchSheets()
Attribute SynchSheets.VB_Description = "Synchronize Sheets"
Attribute SynchSheets.VB_ProcData.VB_Invoke_Func = "K\n14"
'   Duplicates the active sheet's active cell upperleft cell
'   Across all worksheets
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Dim UserSheet As Worksheet, sht As Worksheet
    Dim TopRow As Long, LeftCol As Integer
    Dim UserSel As String
    
    Application.ScreenUpdating = False

'   Remember the current sheet
    Set UserSheet = ActiveSheet
    
'   Store info from the active sheet
    TopRow = ActiveWindow.ScrollRow
    LeftCol = ActiveWindow.ScrollColumn
    UserSel = ActiveWindow.RangeSelection.Address
    
'   Loop through the worksheets
    For Each sht In ActiveWorkbook.Worksheets
        If sht.Visible Then 'skip hidden sheets
            sht.Activate
            Range(UserSel).Select
            ActiveWindow.ScrollRow = TopRow
            ActiveWindow.ScrollColumn = LeftCol
        End If
    Next sht
    
'   Restore the original position
    UserSheet.Activate
    Application.ScreenUpdating = True
End Sub

