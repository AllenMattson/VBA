Attribute VB_Name = "Module1"
Option Explicit

Sub CreatePivotTable()
Attribute CreatePivotTable.VB_Description = "Macro recorded 12/28/1998 by John Walkenbach"
Attribute CreatePivotTable.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim PTcache As PivotCache
    Dim pt As PivotTable
    
    Application.ScreenUpdating = False
'   Delete PivotSheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("PivotSheet").Delete
    On Error GoTo 0
    
'   Create a Pivot Cache
    Set PTcache = ActiveWorkbook.PivotCaches.Create( _
      SourceType:=xlDatabase, _
      SourceData:=Range("A1").CurrentRegion.Address)

'   Add new worksheet
    Worksheets.Add
    ActiveSheet.Name = "PivotSheet"
    ActiveWindow.DisplayGridlines = False

'   Create the Pivot Table from the Cache
    Set pt = ActiveSheet.PivotTables.Add( _
      PivotCache:=PTcache, _
      TableDestination:=Range("A1"), _
      TableName:="BudgetPivot")
    
    With pt
'       Add fields
        .PivotFields("Category").Orientation = xlPageField
        .PivotFields("Division").Orientation = xlPageField
        .PivotFields("Department").Orientation = xlRowField
        .PivotFields("Month").Orientation = xlColumnField
        .PivotFields("Budget").Orientation = xlDataField
        .PivotFields("Actual").Orientation = xlDataField
        .DataPivotField.Orientation = xlRowField
    
'       Add a calculated field to compute variance
        .CalculatedFields.Add "Variance", "=Budget-Actual"
        .PivotFields("Variance").Orientation = xlDataField
        
'       Specify a number format
        .DataBodyRange.NumberFormat = "0,000"
        
'       Apply a style
        .TableStyle2 = "PivotStyleMedium2"
        
'       Hide Field Headers
        .DisplayFieldCaptions = False
        
'       Change the captions
        .PivotFields("Sum of Budget").Caption = " Budget"
        .PivotFields("Sum of Actual").Caption = " Actual"
        .PivotFields("Sum of Variance").Caption = " Variance"
    End With
End Sub
