Attribute VB_Name = "UniqueItems_Table_Function"
Function GetDistinct(ByVal oTarget As Range) As Variant
Dim vArray As Variant
Dim dicMyDictionary As Object
Dim v As Variant
Set dicMyDictionary = CreateObject("Scripting.Dictionary")
vArray = oTarget
For Each v In vArray
    dicMyDictionary(v) = v
Next

GetDistinct = dicMyDictionary.Items()

End Function
Public Function Tbl2Dic(ByVal oTarget As Range) As Object
Dim dicMyDictionary As Object
Dim lRow As Long
Dim lCols As Long
Set dicMyDictionary = CreateObject("Scripting.Dictionary")
lCols = oTarget.Columns.Count - 1

For lRow = 1 To oTarget.Rows.Count
    With oTarget.Cells(lRow, 1)
        Set dicMyDictionary(.Value) = .Offset(0, 1).Resize(1, lCols)
    End With
Next

Set Tbl2Dic = dicMyDictionary

End Function
Sub Test2()
Dim MyDict As Collection

    Set Range1 = ActiveSheet.Range("B4:B90")
    Set Range2 = ActiveSheet.Range("B4:B90")
    array1 = GetDistinct(Range1)
    array2 = GetDistinct(Range2)
    CommonCount = 0
    For i = LBound(array1) To UBound(array1)
        For j = LBound(array2) To UBound(array2)
            MsgBox array1(i) & vbNewLine & array2(j)
        Next j
    Next i

    
End Sub

