Attribute VB_Name = "Module1"
Option Explicit

Sub AddSlicer()
    Dim oSlicerCache As SlicerCache
    Dim oSlicer As Slicer
    Dim oItem As SlicerItem
    
    Set oSlicerCache = ActiveWorkbook.SlicerCaches.Add( _
        Source:=ActiveSheet.PivotTables(1), _
        SourceField:="WarrYears")

    Set oSlicer = oSlicerCache.Slicers.Add( _
        SlicerDestination:=ActiveSheet, _
        Name:="Warranty Years", Caption:="Warranty Years", _
        Top:=14.6551181102362, Left:=481.034409448819)
    
    oSlicer.NumberOfColumns = 3
    oSlicer.Height = 50
    
     With ActiveWorkbook.SlicerCaches("Slicer_WarrYears")
        For Each oItem In .SlicerItems
            If oItem.Value = "3" Then
                oItem.Selected = True
            Else
                oItem.Selected = False
            End If
        Next
    End With
 End Sub


