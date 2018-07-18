Attribute VB_Name = "Module1"
Option Explicit



Sub CreateNewPivot()
    Dim wksData As Worksheet
    Dim rngData As Range
    Dim wksDest As Worksheet
    Dim pvtTable As PivotTable

    ' Set up object variables
    Set wksData = ThisWorkbook.Worksheets("Source Data")
    Set rngData = wksData.UsedRange
    Set wksDest = ThisWorkbook.Worksheets("Sheet2")

    
    ' Check if PivotTable already exists
    If wksDest.PivotTables.Count > 0 Then
        MsgBox "Worksheet " & wksDest.Name & _
            " already contains a pivot table."
        Exit Sub
    End If


    ' Create a skeleton of a PivotTable
    Set pvtTable = wksData.PivotTableWizard(SourceType:=xlDatabase, _
        SourceData:=rngData, TableDestination:=wksDest.Range("B5"))

    ' Close the PivotTable Field List that appears automatically
    ActiveWorkbook.ShowPivotTableFieldList = False

    ' Add fields to the PivotTable
    With pvtTable
        .PivotFields("Vendor").Orientation = xlRowField
        .PivotFields("Equipment Type").Orientation = xlRowField
        .PivotFields("Warranty Type").Orientation = xlColumnField
        With .PivotFields("Equipment Id")
             .Orientation = xlDataField
             .Function = xlCount
        End With
        .PivotFields("Equipment Id").Orientation = xlPageField
    End With

    ' Autofit columns so all headings are visible
    wksDest.UsedRange.Columns.AutoFit
End Sub





