Attribute VB_Name = "Module1"
Option Explicit

Sub DeferLayoutUpdate()
Attribute DeferLayoutUpdate.VB_ProcData.VB_Invoke_Func = " \n14"
    With Sheet1.PivotTables(1)
       .ManualUpdate = True
       With .CubeFields("[Products].[Category]")
            .Orientation = xlRowField
            .Position = 1
       End With
       .ManualUpdate = False
    End With
End Sub


Sub GetDataModel_Info()
    Dim wkb As Workbook
    Dim tbl As Variant
    
    Set wkb = ActiveWorkbook
    
    Debug.Print "Model Name: " & wkb.Model.Name
    Debug.Print "Relationships: " & wkb.Model.ModelRelationships.Count
    Debug.Print "Number of Tables: " & wkb.Model.ModelTables.Count
    Debug.Print "--TABLE NAMES--"
    For Each tbl In wkb.Model.ModelTables
        Debug.Print tbl.Name
    Next
End Sub



