Attribute VB_Name = "F5"
Option Private Module
Option Explicit
Public Function createGetGAdataURL(profileNumber As Variant, ByVal metrics As String, startDate As Variant, endDate As Variant, Optional ByVal filters As String, Optional ByVal dimensions As String, Optional ByVal segment As String, Optional ByVal sortStr As String, Optional maxResults As Long = 10000, Optional startIndex As Long = 1) As String

    On Error Resume Next

    Dim startDateString As String
    Dim endDateString As String

    Dim errorStr As String

    Dim URL As String

    If startDate > endDate Then
        createGetGAdataURL = "Error: Start date should be before end date"
        Exit Function
    End If



    If IsDate(startDate) Then
        startDateString = Year(startDate) & "-" & Format(Month(startDate), "00") & "-" & Format(Day(startDate), "00")
    Else
        startDateString = startDate
    End If

    If IsDate(endDate) Then
        endDateString = Year(endDate) & "-" & Format(Month(endDate), "00") & "-" & Format(Day(endDate), "00")
    Else
        endDateString = endDate
    End If



    If dimensions <> vbNullString Then
        dimensions = Replace(dimensions, "PageDir1", "pagepath")
        dimensions = Replace(dimensions, "PageDir2", "pagepath")
        dimensions = Replace(dimensions, "PageDir3", "pagepath")
        dimensions = Replace(dimensions, "PageDir4", "pagepath")
        dimensions = Replace(dimensions, "visitLengthCategorized", "visitLength")
    End If




    If allProfilesInOneQuery Then
        URL = URL & "profiles=" & uriEncode(allProfilesStr)
    Else
        URL = URL & "profiles=" & uriEncode(profileNumber)
    End If


    URL = URL & "&start-date=" & uriEncode(startDateString)
    URL = URL & "&end-date=" & uriEncode(endDateString)

    If dimensions <> vbNullString Then URL = URL & "&dimensions=" & uriEncode(dimensions)
    If metrics <> vbNullString Then URL = URL & "&metrics=" & uriEncode(metrics)
    If filters <> vbNullString Then URL = URL & "&filters=" & uriEncode(filters)
    If segment <> vbNullString And segment <> "0" And segment <> "-1" Then URL = URL & "&segment=" & uriEncode(segment & rscL2 & segmentName)
    If sortStr <> vbNullString Then URL = URL & "&sort=" & uriEncode(sortStr)
    If maxResults = 0 Then
        URL = URL & "&maxResultsProfile=10000"
    Else
        URL = URL & "&maxResultsProfile=" & maxResults
    End If
    URL = URL & "&start-index=" & startIndex

    createGetGAdataURL = URL

End Function



Public Function parseResponse(gaResponse As Variant) As Variant

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim TempArray As Variant
    Dim TempArray2 As Variant
    Dim arr1 As Variant
    Dim dataArr As Variant
    Dim dataArrJagged As Variant
    Dim headersArr As Variant
    Dim rowNum As Long
    Dim rowCount As Long
    Dim colNum As Long
    Dim colCount As Long
    Dim noteStr As String
    Dim messageStr As String

    Dim i As Long
    Dim j As Long
    Dim sar As Long
    Dim sar2 As Long
    Dim sar3 As Long
    Dim rivi As Long
    Dim rivi2 As Long
    Dim dimensionsCountCurrentArr As Long
    Dim arrayCompressionRequired As Boolean
    Dim convertRSCL As Boolean

    Dim numericColumnsStart As Long
    Dim numericRowsStart As Long

    Dim licenseResponseStr As String
    Dim arvo As Variant



    arrayCompressionRequired = False


    Dim skipECommerceDimensionModification As Boolean



    Dim levFound As Boolean
    Dim cFolderLevel As Integer
    Dim prevFolderLevel As Integer
    Dim kirjain As Long

    '0 SUCCESS
    '1 notes
    '2 col headers
    '3 data

    ' If SDlabelsQuery = False Then
    metricSetNum = queryArr(queryNum, 19)
    columnModificationsArr = metricSetsArr(metricSetNum, 5)
    'End If

    '    gaResponse = Right(gaResponse, Len(gaResponse) - InStr(1, gaResponse, rscL1) - (Len(rscL1) - 1))
    '    noteStr = Split(gaResponse, rscL1)(0)
    '
    '    If parseVarFromStr(noteStr, "STATUS", rscL3) = "ERROR" Then
    '        ReDim TempArray(1 To 1, 1 To 1)
    '        TempArray(1, 1) = "Error: " & parseVarFromStr(gaResponse, "ERROR", rscL3)
    '        parseResponse = TempArray
    '        Exit Function
    '    End If






    gaResponse = Replace(gaResponse, vbCrLf, "")
    gaResponse = Replace(gaResponse, vbLf, "")
    gaResponse = Replace(gaResponse, vbCr, "")


    noteStr = Split(gaResponse, rscL1)(0)

    If parseVarFromStr(noteStr, "STATUS", rscL2) = "ERROR" Then
        ReDim TempArray(1 To 1, 1 To 1)
        TempArray(1, 1) = "Error: " & parseVarFromStr(gaResponse, "ERROR", rscL2)
        parseResponse = TempArray
        Exit Function
    End If


    numericColumnsStart = CInt(parseVarFromStr(noteStr, "NUMERIC_FORMAT_COLUMNS_START", rscL2))
    numericRowsStart = CInt(parseVarFromStr(noteStr, "NUMERIC_FORMAT_ROWS_START", rscL2))

    showNoteStr = parseVarFromStr(noteStr, "SHOW_NOTE", rscL2)
    messageStr = parseVarFromStr(noteStr, "SHOW_MESSAGE", rscL2)
    doActionStr = parseVarFromStr(noteStr, "DO_ACTION", rscL2)

    If dataSource = "GW" Then
        If dateRangeRestriction = "" Then
            dateRangeRestriction = parseVarFromStr(noteStr, "DRR", rscL2)
        End If
    End If

    If messageStr <> vbNullString Then MsgBox messageStr

    If doActionStr <> vbNullString Then Call doAction(doActionStr)


    If Not SDlabelsQuery And parseVarFromStr(noteStr, "SAMPLED", rscL2) = "IS_SAMPLED" Then reportContainsSampledData = True
    If dataSource = "GA" And SDlabelsQuery = False And IsNumeric(parseVarFromStr(noteStr, "TOTALRESULTS", rscL2)) Then
        If CLng(parseVarFromStr(noteStr, "TOTALRESULTS", rscL2)) > maxResults And SDlabelsQuery = False Then
            givemaxResultsPerQueryWarning = True
        End If
    End If

    dimensionsCountCurrentArr = numericColumnsStart
    If parseVarFromStr(noteStr, "CONVERT_RSCL", rscL2) = "TRUE" Then
        convertRSCL = True
    Else
        convertRSCL = False
    End If



    headersArr = Split(Split(gaResponse, rscL1)(1), rscL2)
    dataArr = Split(Split(gaResponse, rscL1)(2), rscL2)
    On Error Resume Next
    licenseResponseStr = Split(gaResponse, rscL1)(3)
    If debugMode Then On Error GoTo 0

    colCount = UBound(headersArr)
    rowCount = UBound(dataArr)


    If rowCount < 0 Or colCount < 0 Then
        ReDim TempArray(1 To 1, 1 To 1)
        TempArray(1, 1) = "Error: No data found"
        parseResponse = TempArray
        Exit Function
    End If





    ReDim TempArray(0 To rowCount + 1, 1 To colCount + 1)
    ReDim dataArrJ(0 To rowCount)

    For rowNum = 0 To rowCount
        dataArrJ(rowNum) = Split(dataArr(rowNum), rscL3)
    Next rowNum

    Call updateProgressIterationBoxes

    For colNum = 0 To colCount
        TempArray(0, colNum + 1) = headersArr(colNum)
    Next colNum


    Call updateProgressIterationBoxes

    i = 0


    If numericRowsStart > 0 Then
        For rowNum = 0 To numericRowsStart - 1
            For colNum = 0 To colCount
                TempArray(rowNum + 1, colNum + 1) = Left$(dataArrJ(rowNum)(colNum), numberOfCharsThatCanBeReturnedToCell)
                i = i + 1
                If i > 400 Then
                    Call updateProgressIterationBoxes
                    i = 0
                End If
            Next colNum
        Next rowNum
    End If

    If numericColumnsStart > 0 Then
        For rowNum = 0 To rowCount
            For colNum = 0 To numericColumnsStart - 1
                TempArray(rowNum + 1, colNum + 1) = Left$(dataArrJ(rowNum)(colNum), numberOfCharsThatCanBeReturnedToCell)
                i = i + 1
                If i > 400 Then
                    Call updateProgressIterationBoxes
                    i = 0
                End If
            Next colNum
        Next rowNum
    End If

    For rowNum = numericRowsStart To rowCount
        For colNum = numericColumnsStart To colCount
            TempArray(rowNum + 1, colNum + 1) = val(dataArrJ(rowNum)(colNum))
            i = i + 1
            If i > 400 Then
                Call updateProgressIterationBoxes
                i = 0
            End If
        Next colNum
    Next rowNum



    Call updateProgressIterationBoxes



    If storeResultsInSeparateWB = True Then
        Const newSheetForEachQuery As Boolean = False
        If storeWB Is Nothing Then
            Set storeWB = Workbooks.Add
            ThisWorkbook.Activate
            Set storeWBsheet = storeWB.Sheets(1)
            storeWBlastRow = 0
        End If
        If newSheetForEachQuery = True Then
            Set storeWBsheet = storeWB.Sheets.Add
            storeWBlastRow = 0
        End If
        With storeWBsheet

            'store all results in same columns
            '            Debug.Print "Q" & queryNum & " Storing to row " & storeWBlastRow
            '            .Cells(storeWBlastRow + 1, 3).Resize(rowCount + 1, colCount).value = TempArray
            '            .Cells(storeWBlastRow + 1, 1).value = queryNum
            '            .Cells(storeWBlastRow + 2, 1).value = subQueryNum
            '            .Cells(storeWBlastRow + 3, 1).value = profNum
            '            storeWBlastRow = storeWBlastRow + rowCount + 1

            'store each query in separate columns
            .Cells(1, (queryNum - 1) * 6 + 2).Resize(rowCount + 1, colCount).value = TempArray
            .Cells(1, (queryNum - 1) * 6 + 1).value = "query: " & queryNum
            .Cells(2, (queryNum - 1) * 6 + 1).value = "profnum: " & profNum
        End With
    End If

    For sar = 1 To dimensionsCountCurrentArr
        If (SDlabelsQuery = False And InStr(1, columnModificationsStr, "%visitlengthcat->" & sar & "%") > 0) Or (SDlabelsQuery = True And InStr(1, columnModificationsStr, "%visitlengthcatsd->" & sar & "%") > 0) Then
            For rivi = 1 To UBound(TempArray, 1)
                Select Case TempArray(rivi, sar)
                Case 0 To 10
                    TempArray(rivi, sar) = "a. 0-10 seconds"
                Case 11 To 30
                    TempArray(rivi, sar) = "b. 11-30 seconds"
                Case 31 To 60
                    TempArray(rivi, sar) = "c. 31-60 seconds"
                Case 61 To 180
                    TempArray(rivi, sar) = "d. 61-180 seconds"
                Case 181 To 600
                    TempArray(rivi, sar) = "e. 181-600 seconds"
                Case 601 To 1800
                    TempArray(rivi, sar) = "f. 601-1800 seconds"
                Case Is >= 1801
                    TempArray(rivi, sar) = "g. 1801+ seconds"
                End Select
            Next rivi
            arrayCompressionRequired = True
        End If
    Next sar






    If dimensionCountMetricIncluded = True And SDlabelsQuery = False Then
        For i = 1 To UBound(columnModificationsArr)
            sar = columnModificationsArr(i, 1)  'metric col
            sar2 = columnModificationsArr(i, 3)  'dimension col
            If columnModificationsArr(i, 2) = "numberoflandingpages" Or columnModificationsArr(i, 2) = "numberofpages" Then
                For rivi = 1 To UBound(TempArray, 1)
                    If InStr(1, TempArray(rivi, sar2), "?") > 0 Then TempArray(rivi, sar2) = Left(TempArray(rivi, sar2), InStr(1, TempArray(rivi, sar2), "?") - 1)
                    If InStr(1, TempArray(rivi, sar2), "#") > 0 Then TempArray(rivi, sar2) = Left(TempArray(rivi, sar2), InStr(1, TempArray(rivi, sar2), "#") - 1)
                Next rivi
                arrayCompressionRequired = True
            ElseIf columnModificationsArr(i, 2) = "avgdaystotransaction" Or columnModificationsArr(i, 2) = "avgvisitstotransaction" Then
                For rivi = 1 To UBound(TempArray, 1)
                    TempArray(rivi, sar) = TempArray(rivi, sar) * CLng(TempArray(rivi, sar2))
                    TempArray(rivi, sar2) = vbNullString
                Next rivi
                skipECommerceDimensionModification = True
                arrayCompressionRequired = True
                columnModificationsArr(i, 4) = "avg"  'change type to prevent dim value counting
                TempArray = deleteColFromArr(TempArray, sar2)
            End If
        Next i
    End If


    'DaysSinceLastVisit
    'VisitCount
    'PageDepth
    'DaysToTransaction
    'VisitsToTransaction
    'adSlotPosition

    stParam1 = "PGA1.50"

    'add zeroes in front of numeric dimensions to get sorting right
    For sar = 1 To UBound(TempArray, 2)
        If Replace(LCase(TempArray(0, sar)), "ga:", "") = "dayofweek" Then
            stParam1 = "PGA1.53"
            For rivi = 1 To UBound(TempArray, 1)
                Select Case TempArray(rivi, sar)
                Case vbNullString
                Case 0
                    TempArray(rivi, sar) = TempArray(rivi, sar) & " Sunday"
                Case 1
                    TempArray(rivi, sar) = TempArray(rivi, sar) & " Monday"
                Case 2
                    TempArray(rivi, sar) = TempArray(rivi, sar) & " Tuesday"
                Case 3
                    TempArray(rivi, sar) = TempArray(rivi, sar) & " Wednesday"
                Case 4
                    TempArray(rivi, sar) = TempArray(rivi, sar) & " Thursday"
                Case 5
                    TempArray(rivi, sar) = TempArray(rivi, sar) & " Friday"
                Case 6
                    TempArray(rivi, sar) = TempArray(rivi, sar) & " Saturday"
                End Select
            Next rivi
            arrayCompressionRequired = True
        End If
    Next sar



    If arrayCompressionRequired = True Then
        TempArray = compressArray(TempArray, dimensionsCountCurrentArr)
        arrayCompressionRequired = False
    End If





    If dimensionCountMetricIncluded = True And SDlabelsQuery = False Then

        arrayCompressionRequired = True

        For i = 1 To UBound(columnModificationsArr)

            If columnModificationsArr(i, 4) = "dimCountMetric" Then

                sar2 = columnModificationsArr(i, 3)  'dimension col
                sar = columnModificationsArr(i, 1)  'metric col
                If SDothersQuery Then sar = sar - segmDimCount

                For rivi = 1 To UBound(TempArray, 1)
                    Call updateProgressIterationBoxes
                    If TempArray(rivi, sar) >= 1 And TempArray(rivi, sar2) <> "(not set)" And TempArray(rivi, sar2) <> "(direct)" And TempArray(rivi, sar2) <> "(other)" Then
                        TempArray(rivi, sar2) = ""
                        TempArray(rivi, sar) = 1
                    Else
                        TempArray(rivi, sar2) = ""
                        TempArray(rivi, sar) = 0
                    End If
                Next rivi

                TempArray = deleteColFromArr(TempArray, sar2)
                dimensionsCountCurrentArr = dimensionsCountCurrentArr - 1
                arrayCompressionRequired = True
            End If
        Next i

    End If


    stParam1 = "PGA1.60"

    If arrayCompressionRequired = True Then TempArray = compressArray(TempArray, dimensionsCountCurrentArr)


    If convertRSCL Then
        TempArray = arrayReplace(TempArray, "%rscL0%", rscL0, dimensionsCountCurrentArr)
        TempArray = arrayReplace(TempArray, "%rscL1%", rscL1, dimensionsCountCurrentArr)
        TempArray = arrayReplace(TempArray, "%rscL2%", rscL2, dimensionsCountCurrentArr)
        TempArray = arrayReplace(TempArray, "%rscL3%", rscL3, dimensionsCountCurrentArr)
        TempArray = arrayReplace(TempArray, "%rscL4%", rscL4, dimensionsCountCurrentArr)
    End If


    parseResponse = TempArray

End Function

Public Function compressArray(arr As Variant, dimensionColumns As Long) As Variant

    If UBound(arr, 1) >= 1 Then





        Debug.Print "Rows before compressions: " & UBound(arr, 1)

        Dim compressColumns() As Boolean
        ReDim compressColumns(1 To UBound(arr, 2))

        Dim resultArr() As Variant
        ReDim resultArr(0 To UBound(arr, 1), 1 To UBound(arr, 2))

        Dim dimValueArr() As Variant
        ReDim dimValueArr(1 To UBound(arr, 1))

        Dim dimValuesStr As String

        Dim dimValueArrRivi As Long

        Dim maxDimValueArrRivi As Long

        Dim sar As Long
        Dim rivi As Long

        Dim arvo As Variant

        Dim vrivi As Long
        Dim vsar As Long
        Dim vsarDim As Long

        Dim i As Long

        For sar = 1 To UBound(arr, 2)
            resultArr(0, sar) = arr(0, sar)
        Next sar

        For sar = 1 To UBound(arr, 2)
            If sar > dimensionColumns Then
                compressColumns(sar) = True
                Debug.Print "Compressing column " & sar & "  " & arr(0, sar)
            Else
                compressColumns(sar) = False
                vsarDim = sar
                Debug.Print "Not compressing column " & sar & "  " & arr(0, sar)
            End If
        Next sar

        DoEvents


        vrivi = UBound(arr, 1)
        vsar = UBound(arr, 2)

        For rivi = 1 To vrivi

            If usingMacOSX = True Then
                i = i + 1
                If i = 100 Then
                    Call updateProgressAdditionalMessage("Rows processed: " & rivi & "/" & vrivi)
                    i = 0
                End If
            End If

            DoEvents
            dimValuesStr = vbNullString
            For sar = 1 To vsarDim
                dimValuesStr = dimValuesStr & ":" & arr(rivi, sar)
            Next sar


            For dimValueArrRivi = 1 To rivi
                If dimValueArr(dimValueArrRivi) = dimValuesStr Then Exit For
                If dimValueArr(dimValueArrRivi) = "" Then
                    dimValueArr(dimValueArrRivi) = dimValuesStr
                    maxDimValueArrRivi = dimValueArrRivi
                    Exit For
                End If
            Next dimValueArrRivi


            For sar = 1 To vsar
                arvo = arr(rivi, sar)
                If compressColumns(sar) = False Then
                    resultArr(dimValueArrRivi, sar) = arvo
                ElseIf IsNumeric(resultArr(dimValueArrRivi, sar)) = False Or resultArr(dimValueArrRivi, sar) = "" Then
                    If IsNumeric(arvo) = True Then
                        resultArr(dimValueArrRivi, sar) = val(arvo)
                    ElseIf Len(arvo) > 0 Then
                        resultArr(dimValueArrRivi, sar) = arvo
                    End If
                Else
                    resultArr(dimValueArrRivi, sar) = val(resultArr(dimValueArrRivi, sar)) + val(arvo)
                End If
            Next sar

        Next rivi

        DoEvents

        resultArr = removeEmptyRowsFromEndOfArray(resultArr)

        compressArray = resultArr
        Erase resultArr
        Erase dimValueArr

        Debug.Print "Rows after compressions: " & maxDimValueArrRivi

    Else

        compressArray = arr

    End If

End Function

Public Function deleteColFromArr(arr As Variant, colToDelete As Long) As Variant

    Dim tempArr As Variant
    Dim col As Long
    Dim rivi As Long
    Dim targetCol As Long

    ReDim tempArr(LBound(arr, 1) To UBound(arr, 1), LBound(arr, 2) To UBound(arr, 2) - 1)

    targetCol = LBound(arr, 2) - 1
    For col = LBound(arr, 2) To UBound(arr, 2)
        If col <> colToDelete Then
            targetCol = targetCol + 1
            For rivi = LBound(arr, 1) To UBound(arr, 1)
                tempArr(rivi, targetCol) = arr(rivi, col)
            Next rivi
        Else
        End If
    Next col

    deleteColFromArr = tempArr

End Function



Public Function arrayReplace(arr As Variant, oldText As String, newText As String, Optional untilColumn As Long = -1) As Variant

    Dim rivi As Long
    Dim sar As Long

    If untilColumn = -1 Then untilColumn = UBound(arr, 2)
    If untilColumn <= 0 Then
        arrayReplace = arr
        Exit Function
    End If

    For rivi = LBound(arr) To UBound(arr)
        For sar = LBound(arr, 2) To untilColumn
            arr(rivi, sar) = Replace(arr(rivi, sar), oldText, newText)
        Next sar
    Next rivi
    arrayReplace = arr
End Function

Sub doAction(doActionStr As String)
    Select Case doActionStr
    Case "LOGOUT"
        Call logout(False)
        End
    Case "CLEARSAVE"
        dataSource = "GA"
        Call logout(False)
        dataSource = "AW"
        Call logout(False)
        dataSource = "FB"
        Call logout(False)
        dataSource = "AC"
        Call logout(False)
        Call deleteNonCongfigSheets
        ThisWorkbook.Save
        End
    End Select
End Sub
