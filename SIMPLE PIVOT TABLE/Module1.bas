Attribute VB_Name = "Module1"
Option Explicit

Sub RecordedMacro1() 'will not work
    Range("A1").Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create _
        (SourceType:=xlDatabase, _
        SourceData:="Sheet1!R1C1:R13C4", _
        Version:=xlPivotTableVersion14).CreatePivotTable _
        TableDestination:="Sheet5!R3C1", _
        TableName:="PivotTable1", _
        DefaultVersion:=xlPivotTableVersion14
    Sheets("Sheet2").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1") _
        .PivotFields("SalesRep")
            .Orientation = xlRowField
            .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1") _
        .PivotFields("Month")
            .Orientation = xlColumnField
            .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1") _
        .AddDataField ActiveSheet.PivotTables("PivotTable1") _
        .PivotFields("Sales"), "Sum of Sales", xlSum
    With ActiveSheet.PivotTables("PivotTable1"). _
        PivotFields("Region")
            .Orientation = xlPageField
            .Position = 1
    End With
End Sub

Sub CreatePivotTable()
    Dim PTCache As PivotCache
    Dim PT As PivotTable
    
'   Create the cache
    Set PTCache = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=Range("A1").CurrentRegion)
    
'   Add a new sheet for the pivot table
    Worksheets.Add
    
'   Create the pivot table
    Set PT = ActiveSheet.PivotTables.Add( _
        PivotCache:=PTCache, _
        TableDestination:=Range("A3"))
    
'   Add the fields
    With PT
        .PivotFields("Region").Orientation = xlPageField
        .PivotFields("Month").Orientation = xlColumnField
        .PivotFields("SalesRep").Orientation = xlRowField
        .PivotFields("Sales").Orientation = xlDataField
        'no field captions
        .DisplayFieldCaptions = False
    End With
End Sub

