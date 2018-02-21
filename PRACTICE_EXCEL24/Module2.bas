Attribute VB_Name = "Module2"
Option Explicit

Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    ActiveSheet.Shapes.Range(Array("CategoryName")).Select
    ActiveSheet.Shapes("CategoryName").IncrementLeft 0.75
    ActiveSheet.Shapes("CategoryName").IncrementTop 41.25
    With ActiveWorkbook.SlicerCaches("Slicer_CategoryName")
        .SlicerItems("Dairy Products").Selected = True
        .SlicerItems("Beverages").Selected = False
        .SlicerItems("Condiments").Selected = False
        .SlicerItems("Confections").Selected = False
        .SlicerItems("Grains/Cereals").Selected = False
        .SlicerItems("Meat/Poultry").Selected = False
        .SlicerItems("Produce").Selected = False
        .SlicerItems("Seafood").Selected = False
    End With
    ActiveWorkbook.SlicerCaches("Slicer_CategoryName").ClearManualFilter
    With ActiveWorkbook.SlicerCaches("Slicer_CategoryName")
        .SlicerItems("Condiments").Selected = True
        .SlicerItems("Beverages").Selected = False
        .SlicerItems("Confections").Selected = False
        .SlicerItems("Dairy Products").Selected = False
        .SlicerItems("Grains/Cereals").Selected = False
        .SlicerItems("Meat/Poultry").Selected = False
        .SlicerItems("Produce").Selected = False
        .SlicerItems("Seafood").Selected = False
    End With
    With ActiveWorkbook.SlicerCaches("Slicer_CategoryName")
        .SlicerItems("Condiments").Selected = True
        .SlicerItems("Dairy Products").Selected = True
        .SlicerItems("Beverages").Selected = False
        .SlicerItems("Confections").Selected = False
        .SlicerItems("Grains/Cereals").Selected = False
        .SlicerItems("Meat/Poultry").Selected = False
        .SlicerItems("Produce").Selected = False
        .SlicerItems("Seafood").Selected = False
    End With
    ActiveSheet.Shapes("CategoryName").IncrementTop -23.25
End Sub
