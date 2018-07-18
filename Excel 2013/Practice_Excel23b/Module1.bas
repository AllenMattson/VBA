Attribute VB_Name = "Module1"
Option Explicit
Sub PivotWithCalcFields()
    ActiveWorkbook.PivotCaches.Add( _
        SourceType:=xlDatabase, _
        SourceData:="Sheet1!R1C1:R4C4").CreatePivotTable _
        TableDestination:="'[Practice_Excel23b.xlsm]Sheet1'!R4C7", _
        TableName:="Piv1", _
        DefaultVersion:=xlPivotTableVersion10

    With ActiveSheet.PivotTables("Piv1").PivotFields("Product")
        .Orientation = xlRowField
        .Position = 1
    End With

    ActiveSheet.PivotTables("Piv1").AddDataField _
        ActiveSheet.PivotTables("Piv1").PivotFields("2010"), _
        "Sum of 2010", xlSum
    ActiveSheet.PivotTables("Piv1").AddDataField _
        ActiveSheet.PivotTables("Piv1").PivotFields("2009"), _
        "Sum of 2009", xlSum
    ActiveSheet.PivotTables("Piv1").AddDataField _
        ActiveSheet.PivotTables("Piv1").PivotFields("2008"), _
        "Sum of 2008", xlSum
    ActiveSheet.PivotTables("Piv1").CalculatedFields.Add _
        "Change: 2010/2009", "='2010' - '2009'", True
    ActiveSheet.PivotTables("Piv1").CalculatedFields.Add _
        "Change: 2009/2008", "='2009' - '2008'", True
    ActiveSheet.PivotTables("Piv1"). _
        PivotFields("Change: 2010/2009"). _
        Orientation = xlDataField
    ActiveSheet.PivotTables("Piv1"). _
        PivotFields("Change: 2009/2008"). _
        Orientation = xlDataField
'    ActiveSheet.PivotTables("Piv1"). _
     PivotFields("Data").Orientation = xlColumnField

End Sub



