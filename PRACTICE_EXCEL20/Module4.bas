Attribute VB_Name = "Module4"
Option Explicit


Sub HighlightAltRows()
Attribute HighlightAltRows.VB_ProcData.VB_Invoke_Func = " \n14"
   With ActiveSheet.UsedRange
    .FormatConditions.Add Type:=xlExpression, _
        Formula1:="=MOD(Row(),3)=0"
    .FormatConditions(1).Interior.ColorIndex = 7
   End With
End Sub

Sub FormatIconSet()
    If Selection.FormatConditions(1).Type = 4 Then
        MsgBox "This selection is formatted with " & _
        "DataBar conditional format."
    End If
End Sub

Sub AddCanShape()
Attribute AddCanShape.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim oShape As Shape
    
    Set oShape = ActiveSheet.Shapes.AddShape( _
        msoShapeCan, 54, 0, 54, 110)
    With oShape
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorAccent4
        .Fill.Transparency = 0.5
        .Line.Visible = msoFalse
    End With
    Set oShape = Nothing
End Sub


Sub RemoveShapes()
    Dim oShape As Shape
    Dim strShapeName As String
    With ActiveSheet
       For Each oShape In .Shapes
            strShapeName = oShape.Name
            oShape.Delete
            Debug.Print "The Shape Object named " _
            & strShapeName & " was deleted."
        Next oShape
    End With
End Sub
