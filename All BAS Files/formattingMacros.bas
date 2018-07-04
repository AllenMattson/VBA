Attribute VB_Name = "formattingMacros"
Option Private Module
Option Explicit
Sub determineMainFont()
    On Error Resume Next
    fontName = "Calibri"
    If fontIsInstalled("Calibri Light") Then
        fontName = "Calibri Light"
    ElseIf fontIsInstalled("Calibri") Then
        fontName = "Calibri"
    ElseIf fontIsInstalled("Helvetica") Then
        fontName = "Helvetica"
    Else
        fontName = "Arial"
    End If
    Range("mainFont").value = fontName
End Sub
Sub changeSort()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Application.ScreenUpdating = False
    Dim sheetID As String
    sheetID = Cells(1, 1).value
    sheetID = findRangeName(Cells(1, 1))
    Dim sortType As String
    Dim sortRange As Range
    Dim erivi As Long
    Dim vrivi As Long
    Dim rowLabelsCol As Long
    Dim secondLabelColumn As Long
    Dim sortingCol As Long

    rowLabelsCol = fetchValue("rowLabelsCol", ActiveSheet)
    If IsNumeric(fetchValue("rowLabelsCol2", ActiveSheet)) Then
        secondLabelColumn = fetchValue("rowLabelsCol2", ActiveSheet)
    Else
        secondLabelColumn = 0
    End If
    sortingCol = fetchValue("sortingCol", ActiveSheet)
    sortType = fetchValue("sortType", ActiveSheet)
    Set sortRange = ActiveSheet.Range(fetchValue("sortRange", ActiveSheet))
    erivi = sortRange.Cells(1, 1).row
    vrivi = erivi + sortRange.Rows.Count - 1

    Select Case sortType
    Case "alphabetic"
        Call storeValue("sortType", "alphabetic desc", ActiveSheet)
        ActiveSheet.Shapes(sheetID & "sortButton1").TextFrame.Characters.Text = "Sorted alphabetically (desc)"
        If secondLabelColumn > 0 Then
            sortRange.sort key2:=Cells(erivi, secondLabelColumn), key1:=Cells(erivi, rowLabelsCol), order1:=Excel.XlSortOrder.xlDescending, order2:=Excel.XlSortOrder.xlAscending
        Else
            sortRange.sort key2:=Cells(erivi, sortingCol), key1:=Cells(erivi, rowLabelsCol), order1:=Excel.XlSortOrder.xlDescending, order2:=Excel.XlSortOrder.xlDescending
        End If
    Case "alphabetic desc"
        Call storeValue("sortType", "metric desc", ActiveSheet)
        ActiveSheet.Shapes(sheetID & "sortButton1").TextFrame.Characters.Text = "Sorted by 1st metric (desc)"
        sortRange.sort key1:=Cells(erivi, sortingCol), key2:=Cells(erivi, rowLabelsCol), order1:=Excel.XlSortOrder.xlDescending, order2:=Excel.XlSortOrder.xlAscending
    Case "metric desc"
        Call storeValue("sortType", "metric asc", ActiveSheet)
        ActiveSheet.Shapes(sheetID & "sortButton1").TextFrame.Characters.Text = "Sorted by 1st metric (asc)"
        sortRange.sort key1:=Cells(erivi, sortingCol), key2:=Cells(erivi, rowLabelsCol), order1:=Excel.XlSortOrder.xlAscending, order2:=Excel.XlSortOrder.xlAscending
    Case Else
        Call storeValue("sortType", "alphabetic", ActiveSheet)
        ActiveSheet.Shapes(sheetID & "sortButton1").TextFrame.Characters.Text = "Sorted alphabetically"
        If secondLabelColumn > 0 Then
            sortRange.sort key2:=Cells(erivi, secondLabelColumn), key1:=Cells(erivi, rowLabelsCol), order1:=Excel.XlSortOrder.xlAscending, order2:=Excel.XlSortOrder.xlAscending
        Else
            sortRange.sort key2:=Cells(erivi, sortingCol), key1:=Cells(erivi, rowLabelsCol), order1:=Excel.XlSortOrder.xlAscending, order2:=Excel.XlSortOrder.xlDescending
        End If
    End Select
    Application.ScreenUpdating = True
    'Call setChartCategories(numberOfCategories)

End Sub


Sub changeConditionalFormatting()
    On Error Resume Next
    Application.ScreenUpdating = False
    Call checkOperatingSystem
    Dim col As Integer
    Dim condFormType As String
    Dim invertColoursCols As String
    Dim midPointAtZeroCols As String
    Dim firstDataRow As Long
    Dim lastDataRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim rng As Range
    Dim invertColours As Boolean
    Dim midPointAtZero As Boolean


    condFormType = fetchValue("condFormType", ActiveSheet)
    invertColoursCols = fetchValue("invertColoursCols", ActiveSheet)
    midPointAtZeroCols = fetchValue("midPointAtZeroCols", ActiveSheet)
    firstDataRow = fetchValue("firstDataRow", ActiveSheet)
    lastDataRow = fetchValue("lastDataRow", ActiveSheet)
    firstCol = fetchValue("firstMetricCol", ActiveSheet)
    lastCol = fetchValue("lastMetricCol", ActiveSheet)

    Select Case condFormType
    Case "databars"
        condFormType = "databars_contrast"
    Case "databars_contrast"
        condFormType = "colouring"
    Case "colouring"
        condFormType = "colouring_pos"
    Case "colouring_pos"
        condFormType = "colouring_neg"
    Case "colouring_neg"
        condFormType = "icons"
    Case "icons"
        condFormType = "none"
    Case "none", ""
        condFormType = "databars"
    Case Else
        condFormType = "colouring"
    End Select
    Call storeValue("condFormType", condFormType, ActiveSheet)
    With ActiveSheet
        For col = firstCol To lastCol
            Set rng = .Range(.Cells(firstDataRow, col), .Cells(lastDataRow, col))
            '  If rng.FormatConditions.Count > 0 Then
            If InStr(1, invertColoursCols, "|" & col & "|") > 0 Then
                invertColours = True
            Else
                invertColours = False
            End If
            If InStr(1, midPointAtZeroCols, "|" & col & "|") > 0 Then
                midPointAtZero = True
            Else
                midPointAtZero = False
            End If

            rng.FormatConditions.Delete
            If condFormType <> "none" Then Call applyConditionalFormatting(rng, condFormType, invertColours, midPointAtZero)
            'End If
        Next col
    End With

End Sub


Sub applyConditionalFormatting(formatRange As Range, Optional formatType As String = "databars", Optional invertColours As Variant = False, Optional midPointAtZero As Boolean = False)


    On Error Resume Next
    If debugMode = False Then On Error Resume Next


    'if max value in range is zero then set the formatting maxpoint value to 1
    '    Dim maxValueIsZero As Boolean
    '    If Application.Max(formatRange) = 0 Then
    '    maxValueIsZero = True
    '    Else
    '    maxValueIsZero = False
    '    End If
    '

    Dim maxValueIsMinValue As Boolean
    Dim maxValue As Double
    maxValue = Application.max(formatRange)
    If maxValue = Application.Min(formatRange) Then
        maxValueIsMinValue = True
    Else
        maxValueIsMinValue = False
    End If


    If excelVersion > 11 Then
        With formatRange
            Select Case formatType
            Case "databars", "databars_contrast"
                .FormatConditions.AddDatabar
                With .FormatConditions(1)
                    .ShowValue = True
                    .SetFirstPriority
                    .MinPoint.Modify newtype:=0, newvalue:=0   'newtype:=xlConditionValueNumber, const value 0

                    If maxValueIsMinValue = True Then
                        If maxValue = 0 Then
                            .MaxPoint.Modify newtype:=0, newvalue:=10000
                            .MinPoint.Modify newtype:=0, newvalue:=1
                        Else
                            .MaxPoint.Modify newtype:=0, newvalue:=maxValue + 1
                            .MinPoint.Modify newtype:=0, newvalue:=maxValue - 1
                        End If
                    Else
                        .MaxPoint.Modify newtype:=2   ' newtype:=xlConditionValueHighestValue
                    End If

                    With .BarColor
                        If formatType = "databars_contrast" Then
                            .Color = RGB(0, 124, 200)
                        Else
                            .Color = RGB(200, 200, 200)
                        End If
                        '  .Color = RGB(15, 37, 63)   'dark blue
                        '.Color = RGB(167, 205, 68)   'green
                    End With
                    If excelVersion >= 14 Then
                        On Error Resume Next
                        '   .BarBorder.Type = xlDataBarBorderSolid
                        .BarFillType = 0  'xlDataBarFillSolid
                        If formatType = "databars_contrast" Then
                            .BarColor.Color = RGB(0, 124, 200)
                        Else
                            .BarColor.Color = RGB(216, 216, 216)
                        End If
                        If debugMode Then On Error GoTo 0
                    End If
                End With
            Case "colouring", "colouring_pos", "colouring_neg"
                .FormatConditions.AddColorScale ColorScaleType:=3
                With .FormatConditions(1)
                    .SetFirstPriority
                    With .ColorScaleCriteria(1)
                        If maxValueIsMinValue = True Then
                            .Type = 0
                            .value = maxValue - 1
                        Else
                            .Type = 1  ' xlConditionValueLowestValue
                        End If

                        If invertColours = True Then
                            If formatType <> "colouring_neg" Then
                                .FormatColor.Color = RGB(173, 234, 0)
                            Else
                                .FormatColor.Color = RGB(255, 255, 255)
                            End If
                        Else
                            If formatType <> "colouring_pos" Then
                                .FormatColor.Color = RGB(229, 27, 0)
                            Else
                                .FormatColor.Color = RGB(255, 255, 255)
                            End If
                        End If
                    End With
                    With .ColorScaleCriteria(2)
                        If midPointAtZero = True Then
                            .Type = 0  '  xlConditionValueNumber
                            .value = 0
                            .FormatColor.Color = RGB(255, 255, 255)
                        Else
                            .Type = 5   'xlConditionValuePercentile
                            .value = 50
                            .FormatColor.Color = RGB(255, 255, 255)
                        End If
                    End With
                    With .ColorScaleCriteria(3)
                        If maxValueIsMinValue = True Then
                            .Type = 0
                            .value = maxValue + 1
                        Else
                            .Type = 2    'xlConditionValueHighestValue
                        End If

                        If invertColours = True Then
                            If formatType <> "colouring_pos" Then
                                .FormatColor.Color = RGB(229, 27, 0)
                            Else
                                .FormatColor.Color = RGB(255, 255, 255)
                            End If
                        Else
                            If formatType <> "colouring_neg" Then
                                .FormatColor.Color = RGB(173, 234, 0)
                            Else
                                .FormatColor.Color = RGB(255, 255, 255)
                            End If
                        End If
                    End With
                End With
            Case "icons"
                .FormatConditions.AddIconSetCondition
                With .FormatConditions(1)
                    .SetFirstPriority
                    .ReverseOrder = False
                    .ShowIconOnly = False
                    .iconSet = ActiveWorkbook.IconSets(16)  'xl5CRV
                    On Error Resume Next
                    .iconSet = ActiveWorkbook.IconSets(20)  'xl5Boxes
                    If debugMode Then On Error GoTo 0
                    With .IconCriteria(2)
                        .Type = 3    ' xlConditionValuePercent
                        .value = 20
                        .Operator = 7
                    End With
                    With .IconCriteria(3)
                        .Type = 3    ' xlConditionValuePercent
                        .value = 40
                        .Operator = 7
                    End With
                    With .IconCriteria(4)
                        .Type = 3    '  xlConditionValuePercent
                        .value = 60
                        .Operator = 7
                    End With
                    With .IconCriteria(5)
                        .Type = 3    'xlConditionValuePercent
                        .value = 80
                        .Operator = 7
                    End With
                End With

            End Select

        End With


    End If


End Sub




