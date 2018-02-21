VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub Workbook_ModelChange( _
    ByVal Changes As ModelChanges)
    
    Dim colTblNames As ModelTableNames
    Dim tblCount As Long
    Dim i As Integer
    
    Set colTblNames = Changes.TablesAdded
    tblCount = colTblNames.Count
    If tblCount > 0 Then
        Debug.Print tblCount & " tables were added."
    Else
        Debug.Print "There are no new tables in the data model."
    End If
    For i = 1 To tblCount
        Debug.Print colTblNames.Item(i)
    Next i
End Sub




