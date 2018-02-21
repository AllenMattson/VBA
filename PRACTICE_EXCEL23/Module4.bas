Attribute VB_Name = "Module4"
Option Explicit

Sub FormatPivotTable()
    Dim pvtTable As PivotTable
    Dim strPiv As String

    If ActiveSheet.PivotTables.Count > 0 Then
        strPiv = ActiveSheet.PivotTables(1).Name
            Set pvtTable = ActiveSheet.PivotTables(strPiv)
    Else
        Exit Sub
    End If

    With pvtTable
        .PivotFields("OrderDate").Orientation = xlRows
        .PivotFields("CompanyName").Orientation = xlHidden
        ' use this statement to group OrderDate by year
        .PivotFields("OrderDate").DataRange.Cells(1).Group _
            Start:=True, End:=True, _
            periods:=Array(False, False, False, False, False, _
            False, True)

        ' use this statement to group OrderDate both by quarter
        ' and year
        ' .PivotFields("OrderDate").DataRange.Cells(1).Group _
            Start:=True, End:=True, _
            periods:=Array(False, False, False, False, _
            False, True, True)

        .PivotFields("OrderDate").Orientation = xlColumns
        .TableRange1.AutoFormat Format:=xlRangeAutoFormatColor2
        .PivotFields("ProductName").DataRange.Select

        ' sort the Product Name field in descending order based
            ' on the
        ' Sum of Total
        .PivotFields("ProductName").AutoSort xlDescending, _
           "Sum of Total """
        Selection.IndentLevel = 2
        With Selection.Font
            .Name = "Times New Roman"
            .FontStyle = "Bold"
            .Size = 10
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    End With
End Sub

Sub Hide1996Data()
    Dim myPivot As PivotTable
    Dim myItem As PivotItem
    Dim strFieldLabel As String

    strFieldLabel = "1996"

    Set myPivot = ActiveSheet.PivotTables(1)
    For Each myItem In myPivot.PivotFields("OrderDate").PivotItems
        If myItem.Name <> strFieldLabel Then
            myItem.Visible = True
        Else
            myItem.Visible = False
        End If
    Next
End Sub


