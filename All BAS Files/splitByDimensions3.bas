Attribute VB_Name = "splitByDimensions3"
Option Private Module
Option Explicit


Dim segmDimNum As Long


Dim segmDimValueStr As String

Dim exitArrRiviLoop As Boolean


Dim metricNumResultArr As Long


Dim dimensionNumResultArr As Long
Dim dimensionNum As Long

Dim profName As String
Dim accountName As String


Dim arvo As Variant

Dim div As Variant
Dim num As Variant

Dim col As Long

Dim dataSar As Long
Dim dataRivi As Long

Dim prevDatarivi As Long

Dim i As Long

Dim rivi As Long
Dim sar As Long

Dim refreshCount As Long

Dim segmDimCategoriesStr As String
Dim maxDimensionStringLength As Long
Dim tempStr As String
Dim placeEachResultOnNewRow As Boolean




Sub fetchFigureSplitByDimensionsLoop()

    On Error GoTo generalErrHandler
    If debugMode = True Then On Error GoTo 0

    Dim skipNext As Boolean
    Dim placeToSheet As Boolean

    Dim dimensionsCountThisQuery As Long
    Dim dataSarOtherColumn As Long

    Dim dimensionsCombinedStrArr() As String
    ReDim dimensionsCombinedStrArr(resultStartRow To resultStartRow)

    If debugMode = False Then Application.Cursor = xlWait

    If queryType = "D" And Not dimensionCountMetricIncluded Then placeEachResultOnNewRow = True


    Dim progressText As String
    Dim progressSmallText As String


    With dataSheet

        stParam1 = "8.15"
        processStatusTimer = Timer
        inDataFetchLoop = True

        Do

            DoEvents

            If Timer - processStatusTimer >= 2 Or objHTTPstatusRunning Then
                Call updateProgressIterationBoxes
                If Not (usingMacOSX Or useQTforDataFetch) Then Call getProcessStatus
                If processQueriesTotal > 0 Then
                    progresspct = 10 + 20 * queriesCompletedCount / queryCount + 50 * processQueriesCompleted / processQueriesTotal
                    '  Debug.Print "COMPL: " & processQueriesCompleted & " TOTAL: " & processQueriesTotal
                    If processQueriesCompleted = 0 Then
                        If processQueriesTotal = 1 Then
                            Call updateProgress(progresspct, "Fetching & processing data...", "Waiting for " & processQueriesTotal & " query to complete", True)
                        Else
                            Call updateProgress(progresspct, "Fetching & processing data...", "Waiting for " & processQueriesTotal & " queries to complete", True)
                        End If
                    ElseIf processQueriesCompleted < processQueriesTotal Then
                        If processQueriesCompleted = 1 Then
                            Call updateProgress(progresspct, "Fetching & processing data...", processQueriesCompleted & " query completed, waiting for " & processQueriesTotal - processQueriesCompleted, True)
                        Else
                            Call updateProgress(progresspct, "Fetching & processing data...", processQueriesCompleted & " queries completed, waiting for " & processQueriesTotal - processQueriesCompleted, True)
                        End If
                    Else
                        Call updateProgress(progresspct, "Fetching & processing data...", "Processing data", True)
                    End If
                    processStatusTimer = Timer
                End If
            End If

            Call updateProgressIterationBoxes


            '  If debugMode = True Then Debug.Print "New iteration..."

            stParam1 = "8.16"

            DoEvents


            'checks for completed queries, stores results and frees up objhttp
            If allQueriesFetched = False Then Call checkForCompletedObjHTTPs

            DoEvents

            stParam1 = "8.161"

            'checks for free objhttps and runs queries
            If allQueriesStarted = False Then Call runQueriesOnFreeObjHTTPs

            DoEvents

            stParam1 = "8.162"

            'parse results
            For queryNum = 1 To queryCount
                If queryArr(queryNum, 9) = True And queryArr(queryNum, 11) = False Then

                    DoEvents

                    stParam1 = "8.1622"
                    If IsArray(arr) Then Erase arr
                    stParam1 = "8.1623"

                    profNum = queryArr(queryNum, 1)
                    iterationNum = queryArr(queryNum, 4)
                    SDlabelsQuery = queryArr(queryNum, 3)
                    SDothersQuery = queryArr(queryNum, 22)

                    skipNext = False
                    '                    If SDlabelsQuery = False And subQueryNum > 1 And dimensionCountMetricIncluded = True Then
                    '                        If queryArr(queryArr(queryNum, 17), 11) = False Then
                    '                            skipNext = True
                    '                        End If
                    '                    End If


                    If Left(queryArr(queryNum, 8), Len("Error:")) = "Error:" Or InStr(1, LCase(queryArr(queryNum, 8)), "<title>500 internal server error</title>") > 0 Then
                        stParam1 = "8.16231"
                        queryArr(queryNum, 11) = True
                        ReDim arr(1 To 1, 1 To 1)
                        arr(1, 1) = queryArr(queryNum, 8)
                        Call checkArrForErrors
                    ElseIf skipNext = False Then
                        If debugMode = True Then Debug.Print "Parsing results for query " & queryNum
                        Call updateProgressAdditionalMessage("Parsing response data")

                        '                        If dataSource = "FB" Then
                        '                            stParam1 = "8.162341"
                        '                            arr = parseFBResponse(queryArr(queryNum, 8))
                        '                        Else
                        stParam1 = "8.16232"
                        arr = parseResponse(queryArr(queryNum, 8))


                        If Not IsArray(arr) Then
                            stParam1 = "8.16235"
                            queryArr(queryNum, 11) = True
                            ReDim arr(1 To 1, 1 To 1)
                            arr(1, 1) = "Error: " & queryArr(queryNum, 8)
                            queryArr(queryNum, 8) = arr
                            Call checkArrForErrors
                        ElseIf UBound(arr, 1) = 1 And Left(arr(1, 1), Len("Error:")) = "Error:" Then
                            stParam1 = "8.16236"
                            queryArr(queryNum, 11) = True
                            queryArr(queryNum, 8) = arr
                            Call checkArrForErrors
                        Else
                            stParam1 = "8.1624"
                            queryArr(queryNum, 8) = arr
                            queryArr(queryNum, 11) = True

                            stParam1 = "8.1625"
                            Call checkArrForErrors
                        End If
                    End If
                End If
            Next queryNum

            stParam1 = "8.163"



            stParam1 = "8.17"

            'checks queries in order, if finished then places data into sheet, exits loop when first unfinished query found

            For queryNum = 1 To queryCount

                If queryArr(queryNum, 11) = True Then
                    If queryArr(queryNum, 10) = False Then
                        iterationNum = queryArr(queryNum, 4)
                        SDlabelsQuery = queryArr(queryNum, 3)
                        SDothersQuery = queryArr(queryNum, 22)
                        placeToSheet = False
                        If queryType = "D" Or SDlabelsQuery = True Then
                            If iterationNum = 1 Or allIteration1queriesPlaced(queryNum) Then placeToSheet = True
                        ElseIf queryType = "SD" Then
                            If queryArr(queryArr(queryNum, 12), 10) And (iterationNum = 1 Or allIteration1queriesPlaced(queryNum)) Then placeToSheet = True  'sdlabels placed
                        End If
                        If placeToSheet Then
                            stParam1 = "8.1701"

                            queryArr(queryNum, 10) = True

                            Debug.Print "Started processing to sheet: " & queryNum
                            stParam1 = "8.1702"
                            Call updateProgressAdditionalMessage("Processing data into sheet")
                            stParam1 = "8.1703"
                            If IsArray(arr) Then Erase arr
                            arr = queryArr(queryNum, 8)
                            queryArr(queryNum, 8) = ""

                            profNum = queryArr(queryNum, 1)
                            metricSetNum = queryArr(queryNum, 19)

                            profID = profilesArr(profNum, 3)
                            profName = profilesArr(profNum, 2)
                            accountName = profilesArr(profNum, 1)
                            email = profilesArr(profNum, 4)




                            If rawDataReport Then
                                segmentNum = 1
                            Else
                                segmentNum = queryArr(queryNum, 21)
                            End If



                            stParam1 = "8.1705"

                            queriesCompletedCount = queriesCompletedCount + 1


                            If SDlabelsQuery = True Then
                                progressSmallText = "Processing data into sheet (categories)"
                            ElseIf iterationNum = 1 Then
                                progressSmallText = "Processing data into sheet"
                            Else
                                progressSmallText = "Processing data into sheet (comparisons to earlier period)"
                            End If

                            If processQueriesTotal > 0 Then
                                progresspct = 10 + 20 * queriesCompletedCount / queryCount + 50 * processQueriesCompleted / processQueriesTotal
                            Else
                                progresspct = 10 + 20 * queriesCompletedCount / queryCount
                            End If
                            Call updateProgress(progresspct, "Fetching & processing data...", progressSmallText)


                            If Not IsArray(arr) Then
                                stParam1 = "8.17051"
                                arvo = arr
                                ReDim arr(1 To 1, 1 To 1)
                                arr(1, 1) = arvo
                            End If

                            stParam1 = "8.1706"

                            'fetch segmenting dimension labels
                            If SDlabelsQuery = True Then

                                stParam1 = "8.171"
                                If IsArray(arr) Then stParam4 = arr(1, 1)

                                ReDim segmDimCategoryArr(1 To segmDimCategoriesCount, 1 To segmDimCount + 1)

                                If Left$(arr(1, 1), 6) <> "Error:" Then

                                    stParam1 = "8.1711"
                                    '(dataSource = "AW" Or dataSource = "AC" Or dataSource = "FB") And
                                    If segmDimIsTime = False Or segmDimHasNonTimeComponent = True Then
                                        If includeOther Then
                                            arr = compressArrayToTopValues(arr, segmDimCount + 1, segmDimCategoriesCount - 1)
                                        Else
                                            arr = compressArrayToTopValues(arr, segmDimCount + 1, segmDimCategoriesCount)
                                        End If
                                    End If
                                    stParam1 = "8.17111"
                                    For segmDimCategoryNum = 1 To segmDimCategoriesCount

                                        If segmDimCategoryNum < segmDimCategoriesCount Or Not includeOther Then
                                            If UBound(arr) < segmDimCategoryNum Then
                                                stParam1 = "8.17112"
                                                segmDimCategoryArr(segmDimCategoryNum, 1) = vbNullString
                                            Else
                                                stParam1 = "8.17113"
                                                segmDimCategoryArr(segmDimCategoryNum, 1) = arr(segmDimCategoryNum, 1)
                                                If segmDimCount > 1 Then
                                                    stParam1 = "8.171131"
                                                    For segmDimNum = 2 To segmDimCount
                                                        segmDimCategoryArr(segmDimCategoryNum, 1) = segmDimCategoryArr(segmDimCategoryNum, 1) & " | " & arr(segmDimCategoryNum, segmDimNum)
                                                    Next segmDimNum
                                                End If
                                                stParam1 = "8.171132"
                                                For segmDimNum = 1 To segmDimCount
                                                    segmDimCategoryArr(segmDimCategoryNum, segmDimNum + 1) = arr(segmDimCategoryNum, segmDimNum)
                                                Next segmDimNum
                                            End If
                                        Else
                                            segmDimCategoryArr(segmDimCategoriesCount, 1) = "Other"
                                        End If


                                        'place sd label to sheet
                                        For iterationNum = 1 To iterationsCount
                                            For metricNum = 1 To metricsCount
                                                dataSar = findColumnNumber()
                                                columnInfoArr(dataSar, 4) = segmDimCategoryArr(segmDimCategoryNum, 1)




                                                If iterationNum = 2 Then
                                                    .Cells(segmDimRow, dataSar).value = "*"
                                                Else
                                                    .Cells(segmDimRow, dataSar).value = segmDimCategoryArr(segmDimCategoryNum, 1)


                                                    If segmDimCount > 1 And segmDimCategoryNum < segmDimCategoriesCount Then
                                                        For segmDimNum = 1 To segmDimCount
                                                            .Cells(segmDimRow + segmDimNum, dataSar).value = segmDimCategoryArr(segmDimCategoryNum, segmDimNum + 1)

                                                        Next segmDimNum
                                                    End If
                                                    If .Cells(segmDimRow, dataSar).value = " | " Then .Cells(segmDimRow, dataSar).value = vbNullString
                                                End If
                                            Next metricNum
                                        Next iterationNum
                                    Next segmDimCategoryNum

                                    stParam1 = "8.17114"
                                    If includeOther Then segmDimCategoryArr(segmDimCategoriesCount, 1) = "Other"
                                ElseIf arr(1, 1) = "Error: No data found" Then

                                    stParam1 = "8.1712"

                                    segmDimCategoryArr(1, 1) = "No data found"
                                    For segmDimCategoryNum = 2 To segmDimCategoriesCount - 1
                                        segmDimCategoryArr(segmDimCategoryNum, 1) = vbNullString
                                    Next segmDimCategoryNum
                                Else

                                    stParam1 = "8.173"

                                    segmDimCategoryArr(1, 1) = arr(1, 1)
                                    For segmDimCategoryNum = 2 To segmDimCategoriesCount - 1
                                        segmDimCategoryArr(segmDimCategoryNum, 1) = vbNullString
                                    Next segmDimCategoryNum

                                End If


                                '      queryArr(queryNum, 8) = segmDimCategoryArr
                                profilesArr(profNum, 5)(segmentNum) = segmDimCategoryArr
                                If debugMode Then Call printArr(segmDimCategoryArr)


                                If IsArray(arr) Then Erase arr

                                Application.ScreenUpdating = True
                                Application.ScreenUpdating = False


                            Else


                                ReDim Preserve dimensionsCombinedStrArr(resultStartRow To vriviData + UBound(arr, 1) + 100)

                                stParam1 = "8.172"

                                ReDim segmDimCategoryArr(1 To 1, 1 To 1)
                                '      If queryType = "SD" Then segmDimCategoryArr = queryArr(queryArr(queryNum, 12), 8)  'fetch segm dim categories from memory
                                If queryType = "SD" Then segmDimCategoryArr = profilesArr(profNum, 5)(segmentNum)    'fetch segm dim categories from memory



                                stParam1 = "8.1721"


                                '                                If metricSetNum = 1 And iterationNum = 1 Then
                                '                                    Call placeColumnHeaders
                                '                                End If


                                stParam1 = "8.173"

                                If Not IsArray(arr) Then
                                    ReDim arr(1 To 1, 1 To 1)
                                    arr(1, 1) = "Error: Data fetch error"
                                End If


                                If Left$(arr(1, 1), 6) <> "Error:" Then


                                    stParam1 = "8.1731"

                                    prevDatarivi = 1
                                    refreshCount = 0

                                    If SDothersQuery Then
                                        dimensionsCountThisQuery = dimensionsCountInclSubGlobal
                                    Else
                                        dimensionsCountThisQuery = dimensionsCountInclSubGlobal + segmDimCount
                                    End If


                                    For arrRivi = 1 To UBound(arr)

                                        'Call addToTimer(1, "End")
                                        If arrRivi Mod 200 = 0 Then Call updateProgressIterationBoxes("")
                                        If arrRivi Mod 400 = 0 Then Call updateProgress(progresspct, "Fetching & processing data...", progressSmallText & " row: " & arrRivi)

                                        stParam1 = "8.1732"
                                        'Call addToTimer(9, "Start0")
                                        '                                        refreshCount = refreshCount + 1
                                        '                                        If refreshCount >= 200 Then
                                        ' doevents
                                        '                                            refreshCount = 0
                                        '                                        End If

                                        'Call addToTimer(10, "Start1")

                                        If arr(arrRivi, 1) = vbNullString Then
                                            exitArrRiviLoop = True
                                            For sar = LBound(arr, 2) To UBound(arr, 2)
                                                If arr(arrRivi, sar) <> vbNullString Then
                                                    exitArrRiviLoop = False
                                                    Exit For
                                                End If
                                            Next sar

                                            If exitArrRiviLoop = True Then Exit For
                                        End If


                                        metricNumResultArr = 1
                                        '     metricNumResultArr = metricSetsArr(metricSetNum, 3) + 1

                                        'Call addToTimer(2, "Start")

                                        stParam1 = "8.1733"

                                        'store segmdim values to arr
                                        If queryType = "SD" Then
                                            For segmDimNum = 1 To segmDimCount
                                                segmDimValuesArr(segmDimNum) = arr(arrRivi, dimensionsCountInclSubGlobal + segmDimNum)
                                            Next segmDimNum
                                        End If


                                        'Call addToTimer(3, "SDlabels")


                                        stParam1 = "8.1734"

                                        Call combineDimensionLabels

                                        'Call addToTimer(4, "combineDimensionLabels")

                                        stParam1 = "8.1741"


                                        dataRivi = 0
                                        If Not placeEachResultOnNewRow Then
                                            'find row where combined dimensionlabels match
                                            dataRivi = findRowWithValue(dimensionsCombinedCol, dimensionsCombined, prevDatarivi, dataSheet, 1, vriviData)
                                            If dataRivi = -1 Then dataRivi = 0
                                            prevDatarivi = dataRivi
                                        End If


                                        'Call addToTimer(5, "findrow")


                                        stParam1 = "8.1742"


                                        If queryType = "SD" Then
                                            If SDothersQuery Then
                                                segmDimCategoryNum = segmDimCategoriesCount
                                            Else
                                                ' If includeOther Then
                                                '     segmDimCategoryNum = segmDimCategoriesCount
                                                ' Else
                                                segmDimCategoryNum = -1
                                                ' End If
                                                segmDimValueStr = segmDimValuesArr(1)
                                                If segmDimCount > 1 Then
                                                    For segmDimNum = 2 To segmDimCount
                                                        segmDimValueStr = segmDimValueStr & " | " & segmDimValuesArr(segmDimNum)
                                                    Next segmDimNum
                                                End If
                                                For i = 1 To UBound(segmDimCategoryArr, 1)
                                                    If segmDimCategoryArr(i, 1) = segmDimValueStr Then
                                                        segmDimCategoryNum = i
                                                        Exit For
                                                    End If
                                                Next i
                                            End If
                                        Else
                                            segmDimCategoryNum = 1
                                        End If

                                        'Call addToTimer(6, "SD2")


                                        stParam1 = "8.1743"
                                        If segmDimCategoryNum <> -1 Then
                                            If iterationNum = 1 Then

                                                If dataRivi = 0 Then

                                                    dataRivi = vriviData + 1

                                                    dimensionsCombinedStrArr(dataRivi) = dimensionsCombined

                                                    If vriviData < rowLimit Then
                                                        dimensionNumResultArr = 1
                                                        'mark dimensionlabels to new row when row not found
                                                        For dimensionNum = 1 To dimensionsCount
                                                            col = resultStartColumn + dimensionNum - 1
                                                            '    If rawDataReport Then col = col + 1
                                                            If dimensionsArr(dimensionNum, 4) >= 2 Then

                                                                .Cells(dataRivi, col).value = arr(arrRivi, dimensionNumResultArr) & "|" & arr(arrRivi, dimensionNumResultArr + 1)

                                                            Else
                                                                If Left(arr(arrRivi, dimensionNumResultArr), 1) = "=" Or Left(arr(arrRivi, dimensionNumResultArr), 1) = "=" Or Left(arr(arrRivi, dimensionNumResultArr), 1) = "=" Then
                                                                    arr(arrRivi, dimensionNumResultArr) = "'" & arr(arrRivi, dimensionNumResultArr)
                                                                End If
                                                                .Cells(dataRivi, col).value = arr(arrRivi, dimensionNumResultArr)

                                                            End If
                                                            dimensionNumResultArr = dimensionNumResultArr + dimensionsArr(dimensionNum, 4)
                                                        Next dimensionNum


                                                        stParam1 = "8.174"

                                                        'shorten very long dimension strings
                                                        If Len(dimensionsCombined) > numberOfCharsThatCanBeReturnedToCell Then

                                                            Debug.Print "Max dimension length " & numberOfCharsThatCanBeReturnedToCell & " exceeded: " & dimensionsCombined; ""

                                                            maxDimensionStringLength = Round((numberOfCharsThatCanBeReturnedToCell - 10 - 3 * (dimensionsCount - 1) / dimensionsCount), 0)
                                                            maxDimensionStringLength = Round(Evaluate("(" & "245-3" & "*" & "(" & dimensionsCount & "-1" & ")" & ")" & "/" & dimensionsCount), 0)
                                                            dimensionNumResultArr = 1

                                                            For dimensionNum = 1 To dimensionsCount
                                                                If dimensionsArr(dimensionNum, 4) = 2 Then
                                                                    If segmDimIncludesYear And (LCase(dimensionsArr(dimensionNum, 2)) = "month" Or LCase(dimensionsArr(dimensionNum, 2)) = "week" Or LCase(dimensionsArr(dimensionNum, 2)) = "weekiso") Then
                                                                        'strip out year as it is in the segmdim
                                                                        arvo = arr(arrRivi, dimensionNumResultArr + 1)
                                                                    Else
                                                                        arvo = arr(arrRivi, dimensionNumResultArr) & "|" & arr(arrRivi, dimensionNumResultArr + 1)
                                                                    End If
                                                                Else
                                                                    If segmDimIncludesYear And (LCase(dimensionsArr(dimensionNum, 2)) = "month" Or LCase(dimensionsArr(dimensionNum, 2)) = "week" Or LCase(dimensionsArr(dimensionNum, 2)) = "weekiso") Then
                                                                        arvo = Right(arr(arrRivi, dimensionNumResultArr), 2)
                                                                    Else
                                                                        arvo = arr(arrRivi, dimensionNumResultArr)
                                                                    End If
                                                                End If
                                                                If dimensionNum = 1 Then
                                                                    dimensionsCombined = Left$(arvo, maxDimensionStringLength)
                                                                Else
                                                                    dimensionsCombined = dimensionsCombined & " | " & Left$(arvo, maxDimensionStringLength)
                                                                End If
                                                            Next dimensionNum

                                                        End If


                                                        .Cells(dataRivi, dimensionsCombinedCol).value = dimensionsCombined



                                                        vriviData = vriviData + 1

                                                    End If

                                                    'Call addToTimer(7, "mark dims")


                                                ElseIf postConcatDimensionIncluded = True Then

                                                    dimensionNumResultArr = 1
                                                    'mark dimensionlabels to new row when row not found
                                                    For dimensionNum = 1 To dimensionsCount
                                                        col = resultStartColumn + dimensionNum - 1

                                                        dimensionNumResultArr = dimensionNumResultArr + dimensionsArr(dimensionNum, 4)
                                                    Next dimensionNum

                                                End If

                                            End If

                                            stParam1 = "8.1744"




                                            If dataRivi <> 0 Then

                                                For metricNum = metricSetsArr(metricSetNum, 1) To metricSetsArr(metricSetNum, 2)
                                                    stParam1 = "8.17441"

                                                    dataSar = findColumnNumber()

                                                    If Left$(arr(1, dimensionsCountThisQuery + metricNumResultArr), 6) = "Error:" Then
                                                        If arr(1, dimensionsCountThisQuery + metricNumResultArr) = "Error: No data found" Then
                                                            .Cells(firstHeaderRow - 1, dataSar).value = "No data found"
                                                        Else
                                                            .Cells(firstHeaderRow - 1, dataSar).value = arr(1, dimensionsCountThisQuery + metricNumResultArr)
                                                        End If

                                                        columnInfoArr(dataSar, 10) = arr(1, dimensionsCountThisQuery + metricNumResultArr)
                                                        metricNumResultArr = metricNumResultArr + metricsArr(metricNum, 4)
                                                    Else
                                                        stParam1 = "8.175"
                                                        num = arr(arrRivi, dimensionsCountThisQuery + metricNumResultArr)
                                                        If metricsArr(metricNum, 5) <> vbNullString Then
                                                            If metricsArr(metricNum, 4) > 1 Then div = arr(arrRivi, dimensionsCountThisQuery + metricNumResultArr + 1)
                                                            Select Case metricsArr(metricNum, 5)
                                                            Case "div"

                                                            Case "1000*div"
                                                                num = 1000 * num
                                                            Case "d1000000"    'AW cost, budget
                                                                num = num / 1000000
                                                                div = 0
                                                            Case "div1000"    'AW CPM
                                                                div = 1000 * div
                                                            Case "div1000000"    'AW CPC, cost per conversion
                                                                div = 1000000 * div
                                                            Case "div*86400"
                                                                div = 86400 * div
                                                            Case "d86400"    'time on site
                                                                num = num / 86400
                                                                div = 0
                                                            Case "div*86400&minus"  'avg time on page
                                                                div = 86400 * (arr(arrRivi, dimensionsCountThisQuery + metricNumResultArr + 1) - arr(arrRivi, dimensionsCountThisQuery + metricNumResultArr + 2))
                                                            Case "minus"    'lost impressions

                                                            Case "minus&div"    'net fan growth rate
                                                                num = arr(arrRivi, dimensionsCountThisQuery + metricNumResultArr) - arr(arrRivi, dimensionsCountThisQuery + metricNumResultArr + 1)
                                                                div = arr(arrRivi, dimensionsCountThisQuery + metricNumResultArr + 2)
                                                            Case "div&minus&minusone"    'viral amplification %
                                                                div = arr(arrRivi, dimensionsCountThisQuery + metricNumResultArr + 1) - arr(arrRivi, dimensionsCountThisQuery + metricNumResultArr + 2)
                                                            Case "div&minus&plus&minusone"    'net fan growth rate, accurate formula
                                                                div = arr(arrRivi, dimensionsCountThisQuery + metricNumResultArr + 1) - arr(arrRivi, dimensionsCountThisQuery + metricNumResultArr + 2) + arr(arrRivi, dimensionsCountThisQuery + metricNumResultArr + 3)
                                                            Case Else
                                                                div = 0
                                                            End Select
                                                            metricNumResultArr = metricNumResultArr + metricsArr(metricNum, 4)
                                                        Else
                                                            metricNumResultArr = metricNumResultArr + 1
                                                        End If

                                                        If queryType = "SD" And Not SDothersQuery And segmDimCategoryNum = segmDimCategoriesCount And includeOther Then
                                                            'don't process invisible SD category values, are handled by SDothersQuery
                                                            div = 0
                                                            num = 0
                                                        Else


                                                            If iterationNum = 1 Or segmDimIsTime = False Or segmDimCategoryNum < segmDimCategoriesCount Or Not includeOther Then  'if segmdim has time, then don't compare category other, makes no sense

                                                                stParam1 = "8.17461"
                                                                stParam4 = "r: " & dataRivi & " c: " & dataSar
                                                                With .Cells(dataRivi, dataSar)
                                                                    If .value = vbNullString Then
                                                                        .value = num
                                                                    Else
                                                                        .value = .value + num
                                                                    End If
                                                                End With
                                                                If metricsArr(metricNum, 4) > 1 Then
                                                                    stParam1 = "8.174611"
                                                                    With tempSheet.Cells(dataRivi, dataSar)
                                                                        If .value = vbNullString Then
                                                                            .value = div
                                                                        Else
                                                                            .value = .value + div
                                                                        End If
                                                                    End With
                                                                End If
                                                            End If

                                                            If (queryType = "D" Or SDothersQuery) And iterationNum = 1 And metricNum = 1 Then

                                                                stParam1 = "8.17463"
                                                                'mark SDothersQuery values to sorting col
                                                                If metricNum = 1 And Not rawDataReport Then
                                                                    stParam1 = "8.17464"
                                                                    For i = 1 To metricsArr(metricNum, 4)
                                                                        With .Cells(dataRivi, sortingCol + i - 1)
                                                                            If .value = vbNullString Then
                                                                                .value = arr(arrRivi, dimensionsCountThisQuery + i)
                                                                            Else
                                                                                .value = .value + arr(arrRivi, dimensionsCountThisQuery + i)
                                                                            End If
                                                                        End With
                                                                    Next i
                                                                End If

                                                            End If


                                                            If queryType = "SD" And Not SDothersQuery And segmDimCategoryNum < segmDimCategoriesCount And includeOther Then
                                                                'substract visible SD category values from others col to avoid duplicate counting
                                                                div = -div
                                                                num = -num


                                                                dataSarOtherColumn = findColumnNumber(, segmDimCategoriesCount)

                                                                With .Cells(dataRivi, dataSarOtherColumn)
                                                                    If .value = vbNullString Then
                                                                        .value = num
                                                                    Else
                                                                        .value = .value + num
                                                                    End If
                                                                End With
                                                                If metricsArr(metricNum, 4) > 1 Then
                                                                    stParam1 = "8.174611"
                                                                    With tempSheet.Cells(dataRivi, dataSarOtherColumn)
                                                                        If .value = vbNullString Then
                                                                            .value = div
                                                                        Else
                                                                            .value = .value + div
                                                                        End If
                                                                    End With
                                                                End If

                                                            End If

                                                        End If
                                                    End If
                                                Next metricNum
                                            End If

                                            'Call addToTimer(8, "markmetrics")
                                        End If

                                    Next arrRivi

                                    placeEachResultOnNewRow = False  'only do this for first query that is placed on sheet

                                    If arrRivi >= maxResults And dataSource = "GA" Then
                                        givemaxResultsPerQueryWarning = True
                                    End If

                                ElseIf iterationNum = 1 And metricSetNum = 1 And Not SDothersQuery Then     'mark errors fetching data

                                    stParam1 = "8.175"

                                    For metricNum = 1 To metricsCount

                                        segmDimCategoryNum = 1

                                        dataSar = findColumnNumber()

                                        If arr(1, 1) = "Error: No data found" Then
                                            '           If iterationNum = 1 Then
                                            '            .Cells(firstHeaderRow - 1, dataSar).value = "No data found"
                                            '           Else
                                            .Cells(firstHeaderRow - 1, dataSar).value = vbNullString
                                            '  End If
                                        Else
                                            .Cells(firstHeaderRow - 1, dataSar).value = arr(1, 1)
                                        End If


                                        columnInfoArr(dataSar, 10) = arr(1, 1)

                                    Next metricNum

                                End If

                                If IsArray(arr) Then Erase arr

                                Application.ScreenUpdating = True
                                Application.ScreenUpdating = False

                            End If

                            Call updateProgress(progresspct, "Fetching & processing data...", "")

                            Debug.Print "Finished: " & queryNum

                            ' If queriesCompletedCount >= queryCount And foundNonFinishedQuery = False Then Exit Do

                            If allQueriesFetched = False And Not usingMacOSX And Not useQTforDataFetch Then Exit For

                        End If
                    End If

                End If

            Next queryNum

            If areAllQueriesPlacedOnSheet() = True Then Exit Do

        Loop
        inDataFetchLoop = False

        Application.Cursor = xlNormal

        ' Call displayTimers
        ' End

        If Not useQTforDataFetch Then Set objHTTPstatus = Nothing


        stParam1 = "8.17501"
        stParam4 = vbNullString
        Call eraseObjHTTPs
        If IsArray(arr) Then Erase arr
        If IsArray(segmDimCategoryArr) Then Erase segmDimCategoryArr

        stParam1 = "8.176"


        stParam1 = "8.1761"

        If sendMode = True Then Call checkE(email, dataSource)

        Call updateProgressIterationBoxes("EXITLOOP")

        Call updateProgressAdditionalMessage("")


    End With



    If sendMode = True Then Call checkE(email, dataSource)

    Exit Sub


generalErrHandler:

    stParam2 = "REPORTLOOPERROR " & Err.Number & "|" & Err.Description & "|" & Application.StatusBar
    Debug.Print "REPORTLOOPERROR: " & stParam1 & " " & stParam2

    ' Call checkE(email, dataSource, True)

    If Err.Number = 18 Then
        Call hideProgressBox
        Call removeTempsheet
        End
    End If

    Resume Next

End Sub



Public Function findColumnNumber(Optional str As String = "", Optional segmDimCategoryNumLoc As Long = -1) As Long
    Dim i As Long

    If segmDimCategoryNumLoc = -1 Then segmDimCategoryNumLoc = segmDimCategoryNum

    If str = vbNullString Then str = profNum & "|" & metricNum & "|" & segmentNum & "|" & segmDimCategoryNumLoc & "|" & iterationNum
    For i = 1 To UBound(columnInfoArr, 1)
        If columnInfoArr(i, 13) = str Then
            findColumnNumber = i
            Exit Function
        End If
    Next i
    Debug.Print "Can't find column " & str
    findColumnNumber = 0
End Function

Sub placeColumnHeaders()

    Dim metricNumLoc As Long
    Dim segmDimCategoryNumLoc As Long
    Dim segmentNumLoc As Long
    Dim profNumLoc As Long
    Dim profNameLoc As String
    Dim accountNameLoc As String
    Dim dataSar As Long
    Dim arvo As Variant
    Dim segmDimCategoriesStr As String
    Dim segmDimNum As Long
    Dim iterationNumLoc As Integer
    Dim segmentCountLoc As Long

    If rawDataReport Then
        segmentCountLoc = 1  'don't make own column for each segment
    Else
        segmentCountLoc = segmentCount
    End If

    ' profNumLoc = profNum
    ' segmentNumLoc = segmentNum

    With dataSheet

        For profNumLoc = 1 To profileCount

            For metricNumLoc = 1 To metricsCount

                For segmentNumLoc = 1 To segmentCountLoc

                    For segmDimCategoryNumLoc = 1 To segmDimCategoriesCount

                        For iterationNumLoc = 1 To iterationsCount

                            Call updateProgressIterationBoxes("")

                            dataSar = findColumnNumber(profNumLoc & "|" & metricNumLoc & "|" & segmentNumLoc & "|" & segmDimCategoryNumLoc & "|" & iterationNumLoc)

                            stParam1 = "8.1722"

                            profNameLoc = profilesArr(profNumLoc, 2)
                            profID = profilesArr(profNumLoc, 3)
                            email = profilesArr(profNumLoc, 4)
                            accountNameLoc = profilesArr(profNumLoc, 1)
                            segmentName = segmentArr(segmentNumLoc, 2)

                            If Not sumAllProfiles And Not rawDataReport Then
                                If dataSource <> "GW" Then .Cells(profIDRow, dataSar).value = profID     'CStr(Int((9999999 - 100000 + 1) * Rnd + 100000))

                                If dataSource <> "FB" And dataSource <> "GW" And dataSource <> "TA" Then .Cells(accountNameRow, dataSar).value = Left$(accountNameLoc, 255)
                                .Cells(profNameRow, dataSar).value = Left$(profNameLoc, 255)
                                If doHyperlinks Then
                                    If dataSource = "YT" Then
                                        .Hyperlinks.Add Cells(accountNameRow, dataSar), "http://www.youtube.com/user/" & accountNameLoc    ', "Open channel in browser"
                                        If profID <> "TOTALS" Then
                                            .Hyperlinks.Add Cells(profNameRow, dataSar), "http://www.youtube.com/watch?v=" & profID    ', "Open video in browser"
                                        Else
                                            .Hyperlinks.Add Cells(profNameRow, dataSar), "http://www.youtube.com/user/" & accountNameLoc    ', "Open channel in browser"
                                        End If
                                    End If
                                    '.Cells(accountNameRow, dataSar).Font.Bold = True
                                    '.Cells(profNameRow, dataSar).Font.Bold = True

                                End If
                            End If

                            If segmentCount > 1 And Not rawDataReport Then .Cells(segmentRow, dataSar).value = Left$(segmentName, 255)

                            stParam1 = "8.172201"

                            .Cells(metricNameRow, dataSar).value = metricsArr(metricNumLoc, 1)
                            If dataSource = "GA" And goalsIncluded = True And metricsArr(metricNumLoc, 10) <> "" Then
                                'goal names
                                stParam1 = "8.172202"
                                arvo = getGoalName(profID, metricsArr(metricNumLoc, 10))
                                stParam1 = "8.172203"
                                If arvo <> vbNullString Then
                                    If InStr(1, metricsArr(metricNumLoc, 1), ": " & arvo) = 0 Then
                                        .Cells(metricNameRow, dataSar).value = metricsArr(metricNumLoc, 1) & ": " & arvo
                                        If profileCount = 1 And segmDimCategoryNumLoc = 1 And iterationNumLoc = 1 And arvo <> vbNullString Then
                                            stParam1 = "8.172204"
                                            metricsArr(metricNumLoc, 1) = metricsArr(metricNumLoc, 1) & ": " & arvo
                                        End If
                                    End If
                                End If
                            End If

                            stParam1 = "8.1723"

                            columnInfoArr(dataSar, 1) = metricsArr(metricNumLoc, 1)
                            columnInfoArr(dataSar, 2) = metricsArr(metricNumLoc, 8)
                            columnInfoArr(dataSar, 3) = profNameLoc
                            '  columnInfoArr(dataSar, 4) = segmDimCategoriesStr
                            If iterationNumLoc = 2 Then
                                columnInfoArr(dataSar, 6) = True
                            Else
                                columnInfoArr(dataSar, 6) = False
                            End If
                            If segmDimCategoryNumLoc = segmDimCategoriesCount Then columnInfoArr(dataSar, 7) = True
                            columnInfoArr(dataSar, 8) = metricsArr(metricNumLoc, 2)    'metric code
                            columnInfoArr(dataSar, 9) = metricsArr(metricNumLoc, 4)    'metric submetric count
                            columnInfoArr(dataSar, 11) = metricNumLoc
                            columnInfoArr(dataSar, 12) = profID
                            columnInfoArr(dataSar, 14) = segmentNumLoc


                            stParam1 = "8.1724"
                            If updatingPreviouslyCreatedSheet = False Then
                                With .Cells(1, dataSar).EntireColumn
                                    Select Case metricsArr(metricNumLoc, 6)
                                    Case "%"
                                        .NumberFormat = "0.0 %"
                                    Case "0.00 %"
                                        .NumberFormat = "0.00 %"
                                    Case "h:mm:ss"
                                        .NumberFormat = "h:mm:ss"
                                    Case "[h]:mm"
                                        .NumberFormat = "[h]:mm"
                                    Case "0.0"
                                        .NumberFormat = "0.0"
                                    Case "0.00"
                                        .NumberFormat = "0.00"
                                    Case Else
                                        .NumberFormat = "#,##0"
                                    End Select
                                    .ColumnWidth = 10
                                End With
                            End If


                            stParam1 = "8.1725"

                            If iterationNumLoc = 2 Then
                                stParam1 = "8.1726"
                                If Not rawDataReport Then
                                    If Not sumAllProfiles Then
                                        .Cells(profIDRow, dataSar).value = profID
                                        .Cells(accountNameRow, dataSar).value = Left$(accountNameLoc, 255)
                                        .Cells(profNameRow, dataSar).value = Left$(profNameLoc, 255)
                                    End If

                                    .Cells(metricNameRow, dataSar).value = metricsArr(metricNumLoc, 1)
                                    ' .Cells(firstHeaderRow - 1, dataSar).value = "CHANGE"

                                    If queryType <> "SD" Then
                                        If segmentCount > 1 Then
                                            .Cells(segmentRow, dataSar).value = "*"
                                        ElseIf groupByMetric = False Then
                                            .Cells(metricNameRow, dataSar).value = "*"
                                        Else
                                            .Cells(profNameRow, dataSar).value = "*"
                                        End If
                                    End If

                                Else
                                    Select Case comparisonValueType
                                    Case "perc", "abs"
                                        .Cells(1, dataSar).value = "Change in " & metricsArr(metricNumLoc, 1)
                                    Case "val"
                                        .Cells(1, dataSar).value = "Comparison value (" & metricsArr(metricNumLoc, 1) & ")"
                                    End Select

                                End If



                                With .Cells(1, dataSar).EntireColumn
                                    Select Case comparisonValueType
                                    Case "perc"
                                        .NumberFormat = "0.0 %"    'Range("numFormatChange").NumberFormat
                                    Case Else
                                        .NumberFormat = .Cells(1, 1).Offset(, -1).NumberFormat
                                    End Select
                                    If Not rawDataReport Then
                                        .Font.Size = 9
                                        .ColumnWidth = 7
                                    End If
                                End With

                            End If
                        Next iterationNumLoc
                    Next segmDimCategoryNumLoc
                Next segmentNumLoc
            Next metricNumLoc
        Next profNumLoc

    End With

End Sub

Function allIteration1queriesPlaced(forQueryNum As Long) As Boolean
    Dim queryNumLoc As Long

    For queryNumLoc = 1 To queryCount
        If queryArr(queryNumLoc, 1) = queryArr(forQueryNum, 1) And queryArr(queryNumLoc, 21) = queryArr(forQueryNum, 21) Then
            If queryArr(queryNumLoc, 4) = 1 And Not queryArr(queryNumLoc, 10) Then
                allIteration1queriesPlaced = False
                Exit Function
            End If
        End If
    Next queryNumLoc
    allIteration1queriesPlaced = True
End Function
'1 profnum
'2 profid
'3 is segmenting dim query Bool
'4 iterationnum
'5 subquerynum
'6 is running or completed
'7 objHTTPnum where running
'8 query result xml
'9 query is completed
'10 results placed on sheet
'11 result xml is parsed to arr
'12 querynum of parent SD labels query
'13 additional query parameters (AW&AC: fieldsArr)
'14 additional query parameters (AW&AC: other parameters)
'15 foundDimValuesArr (if subquerynum = 1)
'16 error count
'17 querynum of parent query where subquerynum = 1  (foundDimValuesArr stored there)
'18 username
'19 metricSetNum
'20 queryIDforDB
'21 segmentNum

