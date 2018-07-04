Attribute VB_Name = "splitByDimensions6"
Option Explicit
Option Private Module


Sub fillDataColumnNumbers()

    Dim col As Long
    col = firstMetricCol

    Dim profileCountLoc As Long
    Dim segmDimCategoriesCountLoc As Long
    Dim segmentCountLoc As Long

    If queryType = "A" Then
        profileCountLoc = 1
        segmentCountLoc = 1
        segmDimCategoriesCountLoc = 1
    Else
        profileCountLoc = profileCount
        segmentCountLoc = segmentCount
        segmDimCategoriesCountLoc = segmDimCategoriesCount
    End If

    If rawDataReport Then segmentCountLoc = 1



    If segmDimCategoriesCountLoc = 0 Then segmDimCategoriesCountLoc = 1

    If Not groupByMetric Then
        For profNum = 1 To profileCountLoc
            For metricNum = 1 To metricsCount
                For segmentNum = 1 To segmentCountLoc
                    For segmDimCategoryNum = 1 To segmDimCategoriesCountLoc
                        For iterationNum = 1 To iterationsCount
                            columnInfoArr(col, 13) = profNum & "|" & metricNum & "|" & segmentNum & "|" & segmDimCategoryNum & "|" & iterationNum
                            If debugMode Then Debug.Print "Col " & col & ": " & columnInfoArr(col, 13)


                            'column borders
                            If iterationNum = 1 Then
                                If segmDimCategoriesCountLoc > 1 And (profileCountLoc > 1 Or metricsCount > 1 Or segmentCountLoc > 0) Then
                                    If segmDimCategoryNum = 1 Then columnInfoArr(col, 15) = "L"
                                ElseIf segmentCountLoc > 1 And (profileCountLoc > 1 Or metricsCount > 1) Then
                                    If segmentNum = 1 Then columnInfoArr(col, 15) = "L"
                                ElseIf metricsCount > 1 And profileCountLoc > 1 Then
                                    If metricNum = 1 Then columnInfoArr(col, 15) = "L"
                                End If
                            End If

                            col = col + 1
                        Next iterationNum
                    Next segmDimCategoryNum
                Next segmentNum
            Next metricNum
        Next profNum
    Else
        For metricNum = 1 To metricsCount
            For profNum = 1 To profileCountLoc
                For segmentNum = 1 To segmentCountLoc
                    For segmDimCategoryNum = 1 To segmDimCategoriesCountLoc
                        For iterationNum = 1 To iterationsCount
                            columnInfoArr(col, 13) = profNum & "|" & metricNum & "|" & segmentNum & "|" & segmDimCategoryNum & "|" & iterationNum
                            If debugMode Then Debug.Print "Col " & col & ": " & columnInfoArr(col, 13)

                            'column borders
                            If iterationNum = 1 Then
                                If segmDimCategoriesCountLoc > 1 And (profileCountLoc > 1 Or metricsCount > 1 Or segmentCountLoc > 0) Then
                                    If segmDimCategoryNum = 1 Then columnInfoArr(col, 15) = "L"
                                ElseIf segmentCountLoc > 1 And (profileCountLoc > 1 Or metricsCount > 1) Then
                                    If segmentNum = 1 Then columnInfoArr(col, 15) = "L"
                                ElseIf metricsCount > 1 And profileCountLoc > 1 Then
                                    If profNum = 1 Then columnInfoArr(col, 15) = "L"
                                End If
                            End If

                            col = col + 1
                        Next iterationNum
                    Next segmDimCategoryNum
                Next segmentNum
            Next profNum
        Next metricNum
    End If


End Sub


Sub determineHeaderRows()

    If rawDataReport Then

        profIDRow = 0
        accountNameRow = 0
        metricNameRow = 1
        firstHeaderRow = 1
        lastHeaderRow = 1
        segmDimCategoriesCount = 1
        segmDimCount = 0
        If fieldNameIsOk(Range("segmDimName").value) = True Then segmDimCount = segmDimCount + 1
        If fieldNameIsOk(Range("segmDimName2").value) = True Then segmDimCount = segmDimCount + 1

    ElseIf queryType = "SD" Then

        segmDimCategoriesCount = Range("segmDimCategories").value
        If dataSource = "FL" Then segmDimCategoriesCount = segmDimCategoriesCount + 1
        segmDimCount = 0
        If fieldNameIsOk(Range("segmDimName").value) = True Then segmDimCount = segmDimCount + 1
        If fieldNameIsOk(Range("segmDimName2").value) = True Then segmDimCount = segmDimCount + 1

        ReDim segmDimCategoryArr(1 To segmDimCategoriesCount, 1 To segmDimCount + 1)

        profIDRow = 2
        accountNameRow = 3

        If groupByMetric = True Then
            profNameRow = 5
            metricNameRow = 4
        Else
            profNameRow = 4
            metricNameRow = 5
        End If
        segmDimRow = 6
        firstHeaderRow = 2
        lastHeaderRow = 6
        If segmDimCount > 1 Then lastHeaderRow = 6 + segmDimCount

    Else
        profIDRow = 2
        accountNameRow = 3
        If groupByMetric = True Then
            profNameRow = 5
            metricNameRow = 4
        Else
            profNameRow = 4
            metricNameRow = 5
        End If
        firstHeaderRow = 2
        lastHeaderRow = 5
        segmDimCategoriesCount = 1
        segmDimCount = 0
    End If


    If segmentCount > 1 And Not rawDataReport Then
        If queryType = "SD" Then
            segmentRow = segmDimRow
        Else
            segmentRow = lastHeaderRow + 1
        End If
        lastHeaderRow = lastHeaderRow + 1
        segmDimRow = segmDimRow + 1
    Else
        segmentRow = 0
    End If

    firstHeaderRow = 2
    resultStartRow = lastHeaderRow + 1

End Sub

Sub combineDimensionLabels()

    Dim dimensionNumResultArr As Long
    Dim dimensionNum As Long

    Dim arvo As Variant

    Dim cDay As Long
    Dim cMonth As Long
    Dim cYear As Long
    Dim cWeek As Long


    'combine dimensionlabels
    dimensionNumResultArr = 1
    For dimensionNum = 1 To dimensionsCount
        If dimensionNum = dimensionsCount And dimensionCountMetricIncluded = True Then dimensionsArr(dimensionNum, 4) = subDimensionCountOrigForLastDim
        If segmDimIncludesYear = True And (isTime(dimensionsArr(dimensionNum, 2), "month") Or isTime(dimensionsArr(dimensionNum, 2), "week")) Then
            arvo = Right(arr(arrRivi, dimensionNumResultArr), 2)
        Else
            arvo = arr(arrRivi, dimensionNumResultArr)
        End If
        stParam4 = arvo
        If iterationNum = 2 Then
            If timeDimensionIncluded = True Then

                stParam1 = "8.1735"

                If comparisonType = "yearly" Then
                    If isTime(dimensionsArr(dimensionNum, 2), "year") Then
                        stParam1 = "8.173501"
                        arvo = val(arvo) + 1
                    ElseIf isTime(dimensionsArr(dimensionNum, 2), "month") Then
                        If segmDimIncludesYear Then
                            stParam1 = "8.173502"
                            segmDimValuesArr(segmDimNumForYear) = val(segmDimValuesArr(segmDimNumForYear)) + 1
                        Else
                            stParam1 = "8.173503"
                            arvo = val(Left(arvo, 4)) + 1 & Right(arvo, Len(arvo) - 4)
                        End If
                        '  ElseIf LCase(dimensionsArr(dimensionNum, 2)) = "week" Or LCase(dimensionsArr(dimensionNum, 2)) = "weekiso" Then
                    ElseIf isTime(dimensionsArr(dimensionNum, 2), "week") Then
                        If segmDimIncludesYear Then
                            stParam1 = "8.173504"
                            segmDimValuesArr(segmDimNumForYear) = val(segmDimValuesArr(segmDimNumForYear)) + 1
                        Else
                            stParam1 = "8.173505"
                            arvo = val(Left(arvo, 4)) + 1 & Right(arvo, Len(arvo) - 4)
                        End If
                    ElseIf isTime(dimensionsArr(dimensionNum, 2), "date") Then
                        stParam1 = "8.173506"
                        arvo = val(Left(arvo, 4)) + 1 & Right(arvo, Len(arvo) - 4)
                        If segmDimIncludesDate = True Then segmDimValuesArr(segmDimNumForDate) = val(Left(segmDimValuesArr(segmDimNumForDate), 4)) + 1 & Right(segmDimValuesArr(segmDimNumForDate), Len(segmDimValuesArr(segmDimNumForDate)) - 4)
                        If segmDimIncludesYear Then segmDimValuesArr(segmDimNumForYear) = val(segmDimValuesArr(segmDimNumForYear)) + 1
                    End If
                ElseIf comparisonType = "previous" Then
                    stParam1 = "8.17351"
                    If isTime(dimensionsArr(dimensionNum, 2), mostGranularTimeDimension) Then
                        If isTime(dimensionsArr(dimensionNum, 2), "year") Then
                            arvo = val(arvo) + 1
                        ElseIf isTime(dimensionsArr(dimensionNum, 2), "month") Then
                            cMonth = val(Right(arvo, 2))
                            cYear = val(Left(arvo, 4))
                            cMonth = cMonth + 1
                            If segmDimIncludesMonth = True Then segmDimValuesArr(segmDimNumForMonth) = Format(val(segmDimValuesArr(segmDimNumForMonth)) + 1, "00")
                            If cMonth > 12 Then
                                cMonth = 1
                                If segmDimIncludesMonth = True Then segmDimValuesArr(segmDimNumForMonth) = Format(1, "00")
                                cYear = cYear + 1
                                If segmDimIncludesYear Then segmDimValuesArr(segmDimNumForYear) = val(segmDimValuesArr(segmDimNumForYear)) + 1
                            End If
                            If segmDimIncludesYear Then
                                arvo = CStr(Format(cMonth, "00"))
                            Else
                                arvo = cYear & "|" & CStr(Format(cMonth, "00"))
                            End If
                        ElseIf isTime(dimensionsArr(dimensionNum, 2), "week") Then
                            cWeek = val(Right(arvo, 2))
                            cYear = val(Left(arvo, 4))
                            cWeek = cWeek + 1
                            If segmDimIncludesWeek Then segmDimValuesArr(segmDimNumForWeek) = Format(val(segmDimValuesArr(segmDimNumForWeek)) + 1, "00")
                            If cWeek > 53 Then
                                cWeek = 1
                                If segmDimIncludesWeek Then segmDimValuesArr(segmDimNumForWeek) = Format(1, "00")
                                cYear = cYear + 1
                                If segmDimIncludesYear Then segmDimValuesArr(segmDimNumForYear) = val(segmDimValuesArr(segmDimNumForYear)) + 1
                            End If
                            If segmDimIncludesYear Then
                                arvo = CStr(Format(cWeek, "00"))
                            Else
                                arvo = cYear & "|" & CStr(Format(cWeek, "00"))
                            End If
                        ElseIf isTime(dimensionsArr(dimensionNum, 2), "date") Then
                            cDay = val(Right(arvo, 2))
                            cMonth = val(Mid(arvo, 5, 2))
                            cYear = val(Left(arvo, 4))
                            arvo = DateSerial(cYear, cMonth, cDay) + 1
                            arvo = CStr(Year(arvo) & Format(Month(arvo), "00") & Format(Day(arvo), "00"))
                            If segmDimIncludesDate Then
                                cDay = val(Right(segmDimValuesArr(segmDimNumForDate), 2))
                                cMonth = val(Mid(segmDimValuesArr(segmDimNumForDate), 5, 2))
                                cYear = val(Left(segmDimValuesArr(segmDimNumForDate), 4))
                                segmDimValuesArr(segmDimNumForDate) = DateSerial(cYear, cMonth, cDay) + 1
                                segmDimValuesArr(segmDimNumForDate) = CStr(Year(segmDimValuesArr(segmDimNumForDate)) & Format(Month(segmDimValuesArr(segmDimNumForDate)), "00") & Format(Day(segmDimValuesArr(segmDimNumForDate)), "00"))
                            End If
                        ElseIf isTime(dimensionsArr(dimensionNum, 2), "hour") Then
                            arvo = val(arvo) + 1
                            If arvo > 23 Then arvo = 0
                            arvo = CStr(Format(arvo, "00"))
                        End If

                    End If
                End If
            ElseIf segmDimIsTime = True Then           'NO TIME DIMENSIONS, BUT SEGMDIM INCLUDES TIME

                stParam1 = "8.1736"
                If dimensionNum = 1 Then
                    If comparisonType = "yearly" Then
                        If segmDimIncludesYear Then segmDimValuesArr(segmDimNumForYear) = segmDimValuesArr(segmDimNumForYear) + 1
                    ElseIf comparisonType = "previous" Then
                        If segmDimIncludesDate Then
                            cDay = val(Right(segmDimValuesArr(segmDimNumForDate), 2))
                            cMonth = val(Mid(segmDimValuesArr(segmDimNumForDate), 5, 2))
                            cYear = val(Left(segmDimValuesArr(segmDimNumForDate), 4))
                            segmDimValuesArr(segmDimNumForDate) = DateSerial(cYear, cMonth, cDay) + 1
                            segmDimValuesArr(segmDimNumForDate) = CStr(Year(segmDimValuesArr(segmDimNumForDate)) & Format(Month(segmDimValuesArr(segmDimNumForDate)), "00") & Format(Day(segmDimValuesArr(segmDimNumForDate)), "00"))
                            If segmDimIncludesYear Then segmDimValuesArr(segmDimNumForYear) = cYear
                            If segmDimIncludesWeek Then segmDimValuesArr(segmDimNumForWeek) = ISOWeekNum(DateSerial(cYear, cMonth, cDay) + 1)
                            If segmDimIncludesMonth Then segmDimValuesArr(segmDimNumForMonth) = cMonth
                        ElseIf segmDimIncludesWeek Then
                            segmDimValuesArr(segmDimNumForWeek) = Format(val(segmDimValuesArr(segmDimNumForWeek)) + 1, "00")
                            If segmDimValuesArr(segmDimNumForWeek) > 53 Then
                                segmDimValuesArr(segmDimNumForWeek) = Format(1, "00")
                                If segmDimIncludesYear Then segmDimValuesArr(segmDimNumForYear) = val(segmDimValuesArr(segmDimNumForYear)) + 1
                            End If
                        ElseIf segmDimIncludesMonth Then
                            segmDimValuesArr(segmDimNumForMonth) = Format(val(segmDimValuesArr(segmDimNumForMonth)) + 1, "00")
                            If segmDimValuesArr(segmDimNumForMonth) > 12 Then
                                segmDimValuesArr(segmDimNumForMonth) = Format(1, "00")
                                If segmDimIncludesYear Then segmDimValuesArr(segmDimNumForYear) = val(segmDimValuesArr(segmDimNumForYear)) + 1
                            End If
                        ElseIf segmDimIncludesYear Then
                            segmDimValuesArr(segmDimNumForYear) = val(segmDimValuesArr(segmDimNumForYear)) + 1
                        End If
                    End If
                End If

            End If
        End If
        stParam1 = "8.1737"
        If dimensionNum = 1 Then
            dimensionsCombined = arvo
        Else
            dimensionsCombined = dimensionsCombined & " | " & arvo
        End If
        dimensionNumResultArr = dimensionNumResultArr + dimensionsArr(dimensionNum, 4)

    Next dimensionNum

End Sub

