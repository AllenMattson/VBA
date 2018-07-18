Attribute VB_Name = "Module1"
Option Explicit

Public Const APPNAME As String = "Get A Color Function"
Public ColorValue As Variant

Function GetAColor() As Variant
    UserForm1.Show
    GetAColor = ColorValue
End Function

Sub Test_GetAColor1()
'   Example of using the GetAColor function
'   This sub prompts for a color, then changes
'   the color of a shape named Donut
    Dim UserColor As Long
    UserColor = GetAColor()
    If UserColor <> False Then
        With ActiveSheet.Shapes("Donut")
            .Fill.ForeColor.RGB = UserColor
            .Line.ForeColor.RGB = UserColor
        End With
    End If
End Sub

Sub Test_GetAColor2()
'   Example of using the GetAColor function
'   This sub prompts for a color, then changes
'   the color the selected cells
    Dim UserColor As Long
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range."
        Exit Sub
    End If
    UserColor = GetAColor()
    If UserColor <> False Then Selection.Interior.Color = UserColor
End Sub

