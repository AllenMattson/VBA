VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Global variables
Dim flag As Boolean         ' Boolean variable to indicate whether
                            ' to delete a drill-down worksheet
Dim strPivSheet As String   ' String to hold the name of the sheet
                            ' containing the PivotTable
Dim strDrillSheet As String ' String to hold the name of the drill-down
                            ' sheet
Dim strPivSource As String  ' String to hold the name of the worksheet
                            ' with the PivotTable source data

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    If strPivSheet = "" Then Exit Sub
    If Sh.Name <> strPivSheet Then
    If InStr(1, strPivSource, Sh.Name) = 0 Then
        If MsgBox("Do you want to Delete " & Sh.Name & _
            " from the workbook" & vbCrLf _
            & "upon returning to PivotTable report?", _
            vbYesNo + vbQuestion, _
            "Sheet: Delete or Keep") = vbYes Then
            flag = True
            strDrillSheet = Sh.Name
        Else
          flag = False
          Exit Sub
       End If
     End If
    End If
    If ActiveSheet.Name = strPivSheet And flag = True Then
        Application.DisplayAlerts = False
        Worksheets(strDrillSheet).Delete
        Application.DisplayAlerts = True
        flag = False
    End If
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, _
                ByVal Target As Range, Cancel As Boolean)
    With ActiveSheet
        If .PivotTables.Count > 0 Then
            strPivSource = ActiveSheet.PivotTables(1).SourceData
            If ActiveCell.PivotField.Name <> "" And _
                IsEmpty(Target) Then
                    MsgBox "Selected cell has no data " & _
                  "- cannot drill down."
                Cancel = True
                Exit Sub
            End If
            strPivSheet = ActiveSheet.Name
        End If
    End With
End Sub


