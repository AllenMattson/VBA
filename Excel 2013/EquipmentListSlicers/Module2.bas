Attribute VB_Name = "Module2"
Option Explicit

Sub ListSlicers()
    Dim oSlicerCache As SlicerCache
    Dim oSlicerCaches As SlicerCaches
    Dim oSlicer As Slicer
    Dim cnt As Integer
    
    Set oSlicerCaches = ActiveWorkbook.SlicerCaches
    cnt = oSlicerCaches.Count
    
    If cnt > 0 Then
        For Each oSlicerCache In oSlicerCaches
            Debug.Print "Slicer Cache Index|Name:" & _
                oSlicerCache.Index & "|" & oSlicerCache.Name
            Debug.Print "Source Type: " & oSlicerCache.SourceType
            For Each oSlicer In oSlicerCache.Slicers
                Debug.Print vbTab & "Name:" & oSlicer.Name
                Debug.Print vbTab & "Caption:" & oSlicer.Caption
                Debug.Print vbTab & "Cols:" & oSlicer.NumberOfColumns
                Debug.Print vbTab & "Col Width:" & oSlicer.ColumnWidth
                Debug.Print vbTab & "Height:" & oSlicer.Height
                Debug.Print vbTab & "Top:" & oSlicer.Top
                Debug.Print vbTab & "Left:" & oSlicer.Left
                Debug.Print vbTab & "Style:" & oSlicer.Style
                Debug.Print vbTab & "Cache level:" & _
                    oSlicer.SlicerCache.CrossFilterType
            Next
        Next
    End If
End Sub

Sub MoveSlicers()
    ActiveSheet.Shapes.Range(Array("Vendor", _
        "Equipment Type", "Warranty Type")).Select
    Selection.Cut
    Sheets("Sheet2").Select
    Range("b3").Select
    With ActiveSheet
        .Name = "Slicers"
        .Paste
    End With
    
    'arrange windows
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
        .NewWindow
    End With
    Sheets("Sheet1").Select
    ActiveWorkbook.Windows.Arrange _
        ArrangeStyle:=xlVertical
End Sub

