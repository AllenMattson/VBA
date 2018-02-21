VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub _
    Workbook_SheetPivotTableChangeSync(ByVal Sh As Object, _
    ByVal Target As PivotTable)

    MsgBox "Thanks for working with " & _
            "PivotTable (" & Target.Name & _
            ") on  " & Sh.Name & _
            " worksheet."
End Sub
