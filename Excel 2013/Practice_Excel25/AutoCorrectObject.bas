Attribute VB_Name = "AutoCorrectObject"
Option Explicit

' this procedure generates a list of AutoCorrect entries
Sub Auto_Correct()
    Dim myList As Variant
    Dim i As Integer

    myList = Application.AutoCorrect.ReplacementList
    ActiveSheet.Cells(1, 1).Select
    For i = LBound(myList) To UBound(myList)
        With ActiveCell
            .Offset(0, 0).Value = myList(i, 1)
            .Offset(0, 1).Value = myList(i, 2)
            .Offset(1, 0).Select
        End With
    Next
    ActiveSheet.Columns("A:B").AutoFit
    Cells(1, 1).Select
End Sub

' this procedure adds new worksheet entries to the
' AutoCorrect list
Sub Auto_Correct_Batch_Add()
    Dim myRange As Range
    Dim myList As Variant
    Dim strReplaceWhat As String
    Dim strReplaceWith As String
    Dim i As Integer

    ' prompt user to select data for processing
    ' the Type argument ensures that the return value is
    ' a valid cell reference (a Range object).
    Set myRange = Application.InputBox( _
        Prompt:="Highlight the range containing your list", _
        Title:="List Selection", _
        Type:=8)
    If myRange.Columns.Count <> 2 Then Exit Sub

    ' save all the values in the selected range to an array
    myList = myRange.Value

    ' retrieve the values from the array and
    ' add them to the AutoCorrect replacements
    For i = LBound(myList) To UBound(myList)
        strReplaceWhat = myList(i, 1)
        strReplaceWith = myList(i, 2)
        If strReplaceWhat <> "" And strReplaceWith <> "" Then
            Application.AutoCorrect.AddReplacement _
                strReplaceWhat, strReplaceWith
        End If
    Next
End Sub



