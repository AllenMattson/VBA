Attribute VB_Name = "Module1"
Option Explicit

Sub ColorNegative()
'   Makes negative cells red
    Dim cell As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    For Each cell In Selection
        If cell.Value < 0 Then
            cell.Interior.Color = RGB(255, 0, 0)
        Else
            cell.Interior.Color = xlNone
        End If
    Next cell
End Sub

Sub ColorNegative2()
'   Makes negative cells red
    Dim WorkRange As Range
    Dim cell As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    Set WorkRange = Application.Intersect(Selection, ActiveSheet.UsedRange)
    For Each cell In WorkRange
        If cell.Value < 0 Then
            cell.Interior.Color = RGB(255, 0, 0)
        Else
            cell.Interior.Color = xlNone
        End If
    Next cell
End Sub

Sub ColorNegative3()
'   Makes negative cells red
    Dim FormulaCells As Range, ConstantCells As Range
    Dim cell As Range
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    
'   Create subsets of original selection
    On Error Resume Next
    Set FormulaCells = Selection.SpecialCells(xlFormulas, xlNumbers)
    Set ConstantCells = Selection.SpecialCells(xlConstants, xlNumbers)
    On Error GoTo 0

'   Process the formula cells
    If Not FormulaCells Is Nothing Then
        For Each cell In FormulaCells
            If cell.Value < 0 Then
                cell.Interior.Color = RGB(255, 0, 0)
            Else
                cell.Interior.Color = xlNone
            End If
            Next cell
    End If

'   Process the constant cells
    If Not ConstantCells Is Nothing Then
        For Each cell In ConstantCells
            If cell.Value < 0 Then
                cell.Interior.Color = RGB(255, 0, 0)
            Else
                cell.Interior.Color = xlNone
            End If
        Next cell
    End If
End Sub

