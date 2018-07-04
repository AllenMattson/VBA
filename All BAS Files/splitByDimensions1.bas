Attribute VB_Name = "splitByDimensions1"
Option Private Module
Option Explicit

Sub fetchFiguresSplitByDimensions()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Call fetchFiguresSplitByDimensionsInit
    Call fetchFiguresSplitByDimensionsStartQueriesAndFormatSheet
    Call fetchFigureSplitByDimensionsLoop
    Call fetchFigureSplitByDimensionsFormatting
End Sub

Sub fetchFiguresSplitByDimensionsInit()

    Dim startDateRange As Range
    Dim endDateRange As Range
    Dim metricsListStart As Range
    Dim metricsListStartRow As Long
    Dim metricsListStartColumn As Long

    Dim dimensionsListStart As Range
    Dim i As Long
    Dim rivi As Long

    Dim dimensionNum As Long

    Dim segmDimNum As Long

    Dim col As Long
    Dim arvo As Variant
    Dim tempArr As Variant

    Dim tempStr As String
    Dim tempStr2 As String


    Dim timeDimensionsIncludedInclMisc As Boolean

    Dim dimensionCountMetricIncludedInMetricSet As Boolean

    Dim doConditionalFormatting As Boolean
    Dim dimensionsRequiringCompressionIncluded As Boolean
    Dim uniqueCountMetricsIncluded As Boolean

    dimensionsRequiringCompressionInSD = False

    On Error GoTo generalErrHandler
    If debugMode = True Then On Error GoTo 0
    Application.EnableCancelKey = xlErrorHandler


    ReDim columnModificationsArr(1 To 30, 1 To 4)
    '1 column
    '2 metric name
    '3 related column


    Debug.Print ""
    Debug.Print "------------------------------------------------------"
    Debug.Print "------------NEW QUERY D " & Now & "------------"
    Debug.Print "------------------------------------------------------"

    Application.ScreenUpdating = False

    nameEncodingStr = ""
    showNoteStr = ""
    metrics = ""
    dimensions = ""
    dimensionsCount = 0
    metricsCount = 0
    segmDimCount = 0
    metricsCountInclSub = 0
    reportContainsSampledData = False
    processQueriesTotal = 0
    processQueriesCompleted = 0
    processIDsStr = ""
    objHTTPstatusRunning = False
    doHyperlinks = Range("doHyperlinks").value
    If rawDataReport Then doHyperlinks = False


    Set storeWB = Nothing
    storeWBlastRow = 0

    tokenRefreshed = False

    If Range("reportFormattingType").value = 1 Then
        rawDataReport = True
    Else
        rawDataReport = False
    End If

    If Range("sumAllProfiles").value = True Then
        sumAllProfiles = True
    Else
        sumAllProfiles = False
    End If
    If sumAllProfiles Or rawDataReport Then
        allProfilesInOneQuery = True
    Else
        allProfilesInOneQuery = False
    End If

    If Range("includeOther").value <> False Then
        includeOther = True
    Else
        includeOther = False
    End If

    advancedSettings = Range("advancedSettings").value


    Randomize
    Call checkOperatingSystem

    If useQTforDataFetch And Not usingMacOSX And Int((5 * Rnd) + 1) = 1 Then
        Call testConnection  'every fifth run, test if MSXML available
    End If

    Application.EnableEvents = False

    dataSource = Range("dataSource").value    '
    clientLoginModeForGA = Range("clientLoginModeForGA").value

    doTotals = Range("doTotals").value

    Call getProxySettingsIfNeeded

    numberOfCharsThatCanBeReturnedToCell = testNumberOfCharsThatCanBeReturnedToCell()


    stParam1 = "8.01"
    stParam4 = vbNullString


    Call setDatasourceVariables

    fontName = Range("mainFont").value

    dateRangeType = Range("daterangetype").value
    If dateRangeType = vbNullString Or dateRangeType = "custom" Then dateRangeType = "fixed"

    If dateRangeType = "fixed" Then
        Set startDateRange = Range("startDate" & varsuffix)
        Set endDateRange = Range("endDate" & varsuffix)
        If startDateRange.value > endDateRange.value Then
            MsgBox "Invalid date range (start date should be before end date)"
            Call hideProgressBox
            End
        End If

        startDate1 = startDateRange.value
        endDate1 = endDateRange.value

        startDate = startDate1
        endDate = endDate1

    Else
        getDatesForDateRangeType (dateRangeType)
        startDate1 = startDate
        endDate1 = endDate
    End If

    If dataSource = "GA" Then
        If startDate1 < DateSerial(2005, 1, 1) Then startDate1 = DateSerial(2005, 1, 1)
        If endDate1 < DateSerial(2005, 1, 1) Then endDate1 = DateSerial(2005, 1, 1)
    End If


    stParam1 = "8.011"

    comparisonValueType = Range("comparisonValueType").value
    If comparisonValueType = vbNullString Then comparisonValueType = "perc"

    stParam1 = "8.013"

    If debugMode Then Debug.Print "Demo version=" & demoVersion

    If runningSheetRefresh = True Then
        If Range("loggedin" & varsuffix).value <> True Then
            Call hideProgressBox
            If importingFromOldVersion = True Then
                MsgBox "Unable to copy report as you have not logged in with " & serviceName & ". Log in and try again."
            Else
                MsgBox "Unable to refresh report as you have not logged in with " & serviceName & ". Log in and try again."
            End If
            Exit Sub
        End If
    End If



    stParam1 = "8.016"




    stParam1 = "8.017"

    aika1 = Timer


    Application.ScreenUpdating = False


    '    If dataSource = "GW" Then
    '        If Range("loggedin").value = False Then
    '            MsgBox "You need to be logged in to the Google Analytics Module in order to run Webmaster Tools reports"
    '            End
    '        End If
    '    End If

    stParam1 = "8.02"
    queryType = Range("queryType").value

    If sumAllProfiles Then
        groupByMetric = False
    ElseIf Range("grouping").value = "metric" Then
        groupByMetric = True
    Else
        groupByMetric = False
    End If

    createCharts = Range("createCharts").value
    If rawDataReport Then createCharts = False

    If Range("separeteQueryForEachGAMetric").value = True Then
        separeteQueryForEachGAMetric = True
    Else
        separeteQueryForEachGAMetric = False
    End If


    If dataSource = "GA" Then
        If Range("segmentID").value = vbNullString Or Range("segmentID").value = 0 Then
            segmentCount = 1
            ReDim segmentArr(1 To segmentCount, 1 To 2)
            segmentArr(1, 1) = "-1"
            segmentArr(1, 2) = "All Visits"
            segmentIsAllVisits = True
        Else
            segmentIsAllVisits = False
            tempStr = Replace(Range("segmentName").value, ", ", ",")
            segmentID = Range("segmentID").value
            segmentCount = UBound(Split(Range("segmentID").value, ",")) + 1
            ReDim segmentArr(1 To segmentCount, 1 To 2)
            For segmentNum = 1 To segmentCount
                segmentArr(segmentNum, 1) = Split(segmentID, ",")(segmentNum - 1)
                segmentArr(segmentNum, 2) = Split(tempStr, ",")(segmentNum - 1)
            Next segmentNum
        End If
    Else
        segmentCount = 1
        ReDim segmentArr(1 To segmentCount, 1 To 2)
        segmentArr(1, 1) = "-1"
        segmentArr(1, 2) = "All Visits"
        segmentIsAllVisits = True
    End If



    segmDimCount = 0

    Call determineHeaderRows

    If createCharts = True Then
        resultStartColumn = reportStartColumn + 14
    Else
        resultStartColumn = reportStartColumn + 10
    End If

    stParam1 = "8.03"

    If sendMode = True Then Call checkE(email, dataSource)

    If usingMacOSX = False Then
        ProgressBox.Show False
        ProgressBox.tweetLink.Visible = True
    End If

    Call updateProgress(5, "Checking query parameters...", , False)


    stParam1 = "8.04"

    Set metricsListStart = Range("metric1name" & varsuffix)

    metricsListStartRow = metricsListStart.row
    metricsListStartColumn = metricsListStart.Column


    Set dimensionsListStart = Range("dimension1name" & varsuffix)

    deleteEmptyColumns = Range("deleteEmptyColumns").value

    givemaxResultsPerQueryWarning = False
    giveUniqueSumWarning = False

    Set profileListStart = Range("profileListStart" & varsuffix)


    If dataSource = "GA" Then
        maxResults1 = Range("maxresults").value
        maxResults = maxResults1
        maxResultsMultiplierForComparisonQuery = 2
        If maxResults1 <= defaultMaxResultsPerQuery Then
            maxResultsMultiplierForComparisonQuery = 2
        Else
            maxResultsMultiplierForComparisonQuery = 1
        End If
    End If

    avoidSampling = Range("avoidSampling").value

    giveMaxResultsWarning = False

    If Range("conditionalFormattingType").value <> "none" Then doConditionalFormatting = True

    stParam1 = "8.05"


    filterStr = Range("filterstring").value

    Call updateProgress(5, "Checking selected profiles...", , False)

    sheetName = Range("wsname").value

    profNum = 1

    profileCount = Application.CountA(Range("profilesStartCQ").Resize(10000, 1))

    If profileCount = 0 Then
        Application.StatusBar = False
        Call hideProgressBox
        reportRunSuccessful = False
        If runningSheetRefresh = True Then
            MsgBox "Unable to refresh report " & sheetName & ". You no longer have access to the accounts included in this report, or the license of the user that has the access right has expired."
            Exit Sub
        Else
            MsgBox "No " & referToProfilesAs & " have been selected. Select at least one from the list and try again."
            End
        End If
    End If

    ReDim profilesArr(1 To profileCount, 1 To 8)
    i = 0

    rCount = Range("rcount").value

    'get IDs from varssheet, insert to profilesarr
    tempArr = Range("profilesStartCQ").Resize(10000, 1).value
    For rivi = 1 To UBound(tempArr)
        If tempArr(rivi, 1) <> vbNullString Then
            profilesArr(profNum, 3) = tempArr(rivi, 1)
            If profNum = profileCount Then Exit For
            profNum = profNum + 1
        End If
    Next rivi


    Dim found As Boolean
    Dim emailCount As Integer
    With tokensSh
        tempArr = .Range(.Cells(1, loginInfoCol), .Cells(vikarivi(.Cells(1, loginInfoCol)), loginInfoCol + 5)).value
    End With
    tempStr = vbNullString    'checked emails
    tempStr2 = vbNullString    'previous email
    ReDim segmentSDlabelsArr(1 To segmentCount)
    For profNum = 1 To profileCount
        found = False
        For rivi = LBound(tempArr) To UBound(tempArr)
            If tempArr(rivi, 1) = "id" & profilesArr(profNum, 3) Then
                found = True
                Exit For
            End If
        Next rivi
        If Not found Then rivi = 0

        If rivi = 0 Then
            If dataSource = "GA" Then
                MsgBox "This report includes a profile which could not be found from the profile list. The probable cause for this is that your account no longer has access to this profile's data. The profile ID in question is " & profilesArr(profNum, 3)
            Else
                MsgBox "This report includes an account which could not be found from the " & serviceName & " account list. The probable cause for this is that your account no longer has access to this account's data. The account ID in question is " & profilesArr(profNum, 3)
            End If
            End
        End If

        email = trimEM(CStr(tempArr(rivi, 3)))

        profilesArr(profNum, 1) = Split(tempArr(rivi, 6), "%%%")(1)   'account name
        profilesArr(profNum, 2) = Split(tempArr(rivi, 6), "%%%")(0)   'prof name
        profilesArr(profNum, 4) = email    'email
        profilesArr(profNum, 5) = segmentSDlabelsArr    'sdLabelsArr
        profilesArr(profNum, 6) = 0    'queries running

        authToken = tempArr(rivi, 2)
        emailLastCheckedOK = tempArr(rivi, 4)
        profilesArr(profNum, 7) = authToken

        If InStr(1, tempStr, "|" & email) = 0 Then
            emailCount = emailCount + 1
            If Int(emailLastCheckedOK) <> Date Or rCount = 0 Then
                Call checkE(email, dataSource, , , True)
                Call storeEmailCheckedDateToSheet(email)
                tempStr = tempStr & "|" & email
                rCount = rCount + 1
                Range("rcount").value = Range("rcount").value + 1
                i = i + 1
            End If
        End If

        profilesArr(profNum, 8) = profilesArr(profNum, 3)     'id
        profilesArr(profNum, 8) = profilesArr(profNum, 8) & rscL2 & convertRSCL(profilesArr(profNum, 2))      'name
        profilesArr(profNum, 8) = profilesArr(profNum, 8) & rscL2    ' & profilesArr(profNum, 7)      'token
        profilesArr(profNum, 8) = profilesArr(profNum, 8) & rscL2 & convertRSCL(profilesArr(profNum, 1))      'account name

    Next profNum
    Erase tempArr
    tempStr = vbNullString





    If allProfilesInOneQuery Then
        Dim tokenNum As Integer
        Dim prevToken As String
        tokenNum = 0

        Dim sccountNum As Integer
        Dim prevAccount As String
        sccountNum = 0

        allProfilesStr = vbNullString
        For profNum = 1 To profileCount
            If allProfilesStr <> vbNullString Then allProfilesStr = allProfilesStr & rscL1
            allProfilesStr = allProfilesStr & profilesArr(profNum, 3)  'id
            If sumAllProfiles And emailCount = 1 Then

            Else
                If sumAllProfiles Then
                    allProfilesStr = allProfilesStr & rscL2
                Else
                    allProfilesStr = allProfilesStr & rscL2 & convertRSCL(profilesArr(profNum, 2))  'name
                End If
                If emailCount = 1 Then
                    allProfilesStr = allProfilesStr & rscL2
                Else
                    If profilesArr(profNum, 7) <> prevToken Then
                        tokenNum = tokenNum + 1
                        nameEncodingStr = nameEncodingStr & rscL1 & "T" & tokenNum & rscL2 & profilesArr(profNum, 7)
                        prevToken = profilesArr(profNum, 7)
                    End If
                    allProfilesStr = allProfilesStr & rscL2 & rscL3 & "T" & tokenNum & rscL3  'token
                    '   allProfilesStr = allProfilesStr & rscL2 & profilesArr(profNum, 7)  'token
                End If
                If sumAllProfiles Then
                    allProfilesStr = allProfilesStr & rscL2
                Else
                    If profilesArr(profNum, 1) <> prevAccount Then
                        sccountNum = sccountNum + 1
                        nameEncodingStr = nameEncodingStr & rscL1 & "A" & sccountNum & rscL2 & convertRSCL(profilesArr(profNum, 1))
                        prevAccount = profilesArr(profNum, 1)
                    End If
                    allProfilesStr = allProfilesStr & rscL2 & rscL3 & "A" & sccountNum & rscL3   'token
                    '     allProfilesStr = allProfilesStr & rscL2 & convertRSCL(profilesArr(profNum, 1))  'account name
                End If
            End If
        Next profNum
        allProfilesStr = allProfilesStr
        profileCount = 1
    End If

    If debugMode Then Debug.Print allProfilesStr
    If debugMode Then Debug.Print nameEncodingStr
    'End

    stParam1 = "8.06"



    Call updateProgress(6, "Checking selected metrics...", , False)

    stParam1 = "8.0601"


    metrics = ""
    dimensions = ""

    goalsIncluded = False

    metricSetNum = 1
    metricSetsCount = 1
    metricNumInMetricSet = 0
    metricNumInclSubInMetricSet = 0
    dimensionCountMetricIncludedInMetricSet = False
    uniqueCountMetricsIncluded = False

    With metricsListStart.Worksheet

        For metricNum = 1 To 12

            If fieldNameIsOk(.Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn).value) = True Then

                stParam1 = "8.061"

                metricsCount = metricsCount + 1
                metricNumInMetricSet = metricNumInMetricSet + 1

                metricsArr(metricsCount, 1) = .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn).value    'name disp
                metricsArr(metricsCount, 2) = .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 1).value    'name
                metricsArr(metricsCount, 3) = .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 2).value    'metrics list

                If .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 3).value = vbNullString Then
                    metricsArr(metricsCount, 4) = 1
                Else
                    metricsArr(metricsCount, 4) = CLng(.Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 3).value)    'metrics count
                End If
                metricsArr(metricsCount, 5) = .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 4).value    'operation
                metricsArr(metricsCount, 6) = .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 5).value    'formatting
                metricsArr(metricsCount, 7) = .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 6).value    'basemetric

                If metricsCount = 1 Then
                    If metricsArr(metricsCount, 4) = 1 Then
                        firstMetricStr = metricsArr(metricsCount, 3)
                    Else
                        firstMetricStr = metricsArr(metricsCount, 7)
                    End If
                End If

                If dataSource = "GA" Then
                    metricsArr(metricsCount, 9) = .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 8).value    'dimension for dimcount metrics
                    metricsArr(metricsCount, 10) = .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 9).value    'goal number for goal name fetch
                    If metricsArr(metricsCount, 9) = 0 Then metricsArr(metricsCount, 9) = vbNullString
                    If metricsArr(metricsCount, 10) <> "" Then goalsIncluded = True
                End If


                stParam1 = "8.062"
                If metricsArr(metricsCount, 7) = vbNullString Then metricsArr(metricsCount, 7) = metricsArr(metricsCount, 3)

                If dataSource = "AW" Or dataSource = "AC" Then
                    metricsArr(metricsCount, 8) = .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 9).value    'invert condform
                Else
                    metricsArr(metricsCount, 8) = .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 7).value    'invert condform
                End If
                If metricsArr(metricsCount, 8) = vbNullString Then metricsArr(metricsCount, 8) = False
                stParam1 = "8.063"

                If dataSource = "AW" Then
                    metricsArr(metricsCount, 12) = .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 7).value  'don't calculate totals
                End If

                metricsArr(metricsCount, 13) = dataSource


                If dataSource = "GA" And metricNum > 1 Then
                    If separeteQueryForEachGAMetric Or (metricNumInclSubInMetricSet + metricsArr(metricsCount, 4) > 8 Or (dimensionCountMetricIncludedInMetricSet = True Or (metricsArr(metricsCount, 9) <> vbNullString And metricNumInMetricSet > 1))) Then
                        dimensionCountMetricIncludedInMetricSet = False
                        metricSetNum = metricSetNum + 1
                        metricSetsCount = metricSetsCount + 1
                        metricNumInMetricSet = 0
                        metricNumInclSubInMetricSet = 0
                    End If
                End If

                metricsArr(metricsCount, 11) = metricSetNum

                metricsCountInclSub = metricsCountInclSub + metricsArr(metricsCount, 4)
                metricNumInclSubInMetricSet = metricNumInclSubInMetricSet + metricsArr(metricsCount, 4)


                If metricsArr(metricsCount, 9) <> vbNullString Then dimensionCountMetricIncludedInMetricSet = True

                If metricsArr(metricsCount, 2) = "Visitors" Or metricsArr(metricsCount, 2) = "Uniquepurchases" Or metricsArr(metricsCount, 2) = "UniquePageviews" Or metricsArr(metricsCount, 2) = "Searchuniques" Or metricsArr(metricsCount, 2) = "uniqueAppviews" Or metricsArr(metricsCount, 2) = "uniqueSocialInteractions" Or InStr(1, metricsArr(metricsCount, 3), "visitors") > 0 Then uniqueCountMetricsIncluded = True




                stParam1 = "8.064"
                If metricsCount = 1 Then
                    metrics = metricsArr(metricsCount, 3)
                Else
                    metrics = metrics & rscL1 & metricsArr(metricsCount, 3)
                End If


                stParam1 = "8.065"
                ' Debug.Print metrics
            End If

        Next metricNum

    End With
    stParam1 = "8.066"
    Debug.Print "Metrics: " & metrics





    If metricsCount = 0 Then
        If dataSource = "MC" Or dataSource = "FA" Then
            createCharts = False  'Some data sources allow dimension-only queries
            rawDataReport = True
        Else
            Application.StatusBar = False
            MsgBox "Choose at least one metric first"
            Call hideProgressBox
            End
        End If
    End If
    stParam1 = "8.067"



    If dataSource = "GA" Then

        ReDim metricSetsArr(1 To metricSetsCount, 1 To 10)
        '1 first metricNum
        '2 last metricNum
        '3 preceding submetric count
        '4 metrics
        '5 columnModificationsArr
        '6 dimensions (differs when dimcountmetrics included)
        '7 sort metric
        '8 datasource
        '9 profID
        '10 contains dimcountmetric

        metricSetsArr(1, 1) = 1
        metricSetsArr(1, 3) = 0
        metricNumInclSubInMetricSet = 0
        metricNumInMetricSet = 0
        metricSetNum = 1
        arvo = vbNullString
        For metricNum = 1 To metricsCount
            If metricSetNum < metricsArr(metricNum, 11) Then
                metricSetsArr(metricSetNum, 2) = metricNum - 1
                '  If metricSetNum > 1 Then arvo = arvo & "&" & firstMetricStr   'add first metric to other metric sets to get sorting right
                metricSetsArr(metricSetNum, 4) = arvo
                arvo = vbNullString
                metricSetsArr(metricSetNum + 1, 1) = metricNum
                metricSetsArr(metricSetNum + 1, 3) = metricNumInclSubInMetricSet
                metricNumInMetricSet = 1
            Else
                metricNumInMetricSet = metricNumInMetricSet + 1
            End If
            metricSetNum = metricsArr(metricNum, 11)
            metricNumInclSubInMetricSet = metricNumInclSubInMetricSet + metricsArr(metricNum, 4)
            If metricNumInMetricSet = 1 Then
                arvo = metricsArr(metricNum, 3)
                If metricsArr(metricNum, 4) > 1 Then
                    metricSetsArr(metricSetNum, 7) = metricsArr(metricNum, 7) & "_desc"
                Else
                    metricSetsArr(metricSetNum, 7) = metricsArr(metricNum, 3) & "_desc"
                End If
                metricSetsArr(metricSetNum, 8) = metricsArr(metricNum, 13)   'datasource
                metricSetsArr(metricSetNum, 9) = metricsArr(metricNum, 14)   'profid
            Else
                arvo = arvo & rscL1 & metricsArr(metricNum, 3)
            End If
            If metricsArr(metricNum, 9) <> vbNullString Then
                metricSetsArr(metricSetNum, 10) = True  'contains dimcountmetric
            Else
                metricSetsArr(metricSetNum, 10) = False
            End If
        Next metricNum
        metricSetsArr(metricSetsCount, 2) = metricsCount
        If metricSetNum > 1 And metricSetsArr(metricSetNum, 8) = dataSource Then arvo = arvo & rscL1 & firstMetricStr   'add first metric to other metric sets to get sorting right (only if same datasource)
        metricSetsArr(metricSetsCount, 4) = arvo
        If debugMode = True Then
            For metricSetNum = 1 To metricSetsCount
                Debug.Print "MetricSetNum " & metricSetNum & ": " & metricSetsArr(metricSetNum, 4)
            Next metricSetNum
        End If
        '   m = 2: Print m & ": ": Print "first: " & metricSetsArr(m, 1): Print "last: " & metricSetsArr(m, 2): Print "preceding count: " & metricSetsArr(m, 3): Print "metrics: " & metricSetsArr(m, 4)
    Else
        ReDim metricSetsArr(1 To 1, 1 To 10)
        metricSetsCount = 1
        metricSetsArr(1, 1) = 1
        metricSetsArr(1, 2) = metricsCount
        metricSetsArr(1, 3) = 0
        metricSetsArr(1, 4) = metrics
    End If



    If goalsIncluded = True Then goalsArr = Range("goals").value

    stParam1 = "8.068"
    Application.Calculation = xlManual




    stParam1 = "8.08"


    Call updateProgress(6, "Checking selected dimensions...", , False)

    segmDimNameDisp = Range("segmDimNameDisp").value
    segmDimName = Range("segmDimName").value
    segmDimName2 = Range("segmDimName2").value

    timeDimensionIncluded = False
    timeDimensionsIncludedInclMisc = False
    nonTimeDimensionIncluded = False
    segmDimIsTime = False

    columnModificationsStr = ""
    folderDimensionIncluded = False

    dimensionsCount = 0
    dimensionsCountInclSubGlobal = 0
    mostGranularTimeDimension = ""
    dateDimensionIncluded = False
    extraDimensionColumns = 0
    postConcatDimensionIncluded = False
    dimensionsRequiringCompressionIncluded = False

    If rawDataReport And Not sumAllProfiles Then

        dimensionsCount = dimensionsCount + 1
        dimensionsCountInclSubGlobal = dimensionsCountInclSubGlobal + 1


        dimensionsArr(dimensionsCount, 1) = capitalizeFirstLetter(referToAccountsAsSing)
        dimensionsArr(dimensionsCount, 2) = "account"
        dimensionsArr(dimensionsCount, 3) = rscL3 & "account" & rscL3
        dimensionsArr(dimensionsCount, 4) = 1

        If dimensions = vbNullString Then
            dimensions = dimensionsArr(dimensionsCount, 3)
        Else
            dimensions = dimensions & rscL1 & dimensionsArr(dimensionsCount, 3)
        End If

        dimensionsBasicStr = dimensions


        dimensionsCount = dimensionsCount + 1
        dimensionsCountInclSubGlobal = dimensionsCountInclSubGlobal + 1

        dimensionsArr(dimensionsCount, 1) = capitalizeFirstLetter(referToProfilesAsSing)
        dimensionsArr(dimensionsCount, 2) = "profile"
        dimensionsArr(dimensionsCount, 3) = rscL3 & "profile" & rscL3
        dimensionsArr(dimensionsCount, 4) = 1

        If dimensions = vbNullString Then
            dimensions = dimensionsArr(dimensionsCount, 3)
        Else
            dimensions = dimensions & rscL1 & dimensionsArr(dimensionsCount, 3)
        End If

        dimensionsBasicStr = dimensions

    ElseIf rawDataReport Then  'sumallprofiles


    End If

    If rawDataReport And segmentCount > 1 Then

        dimensionsCount = dimensionsCount + 1
        dimensionsCountInclSubGlobal = dimensionsCountInclSubGlobal + 1

        dimensionsArr(dimensionsCount, 1) = "Segment"
        dimensionsArr(dimensionsCount, 2) = "segment"
        dimensionsArr(dimensionsCount, 3) = rscL3 & "segment" & rscL3
        dimensionsArr(dimensionsCount, 4) = 1

        If dimensions = vbNullString Then
            dimensions = dimensionsArr(dimensionsCount, 3)
        Else
            dimensions = dimensions & rscL1 & dimensionsArr(dimensionsCount, 3)
        End If

        dimensionsBasicStr = dimensions
    End If


    With dimensionsListStart.Worksheet

        For dimensionNum = 1 To 10


            If fieldNameIsOk(.Cells(dimensionNum + dimensionsListStart.row - 1, dimensionsListStart.Column).value) = True Then

                dimensionsCount = dimensionsCount + 1


                dimensionsArr(dimensionsCount, 1) = .Cells(dimensionNum + dimensionsListStart.row - 1, dimensionsListStart.Column).value    'name disp
                dimensionsArr(dimensionsCount, 2) = .Cells(dimensionNum + dimensionsListStart.row - 1, dimensionsListStart.Column + 1).value    'name


                If dataSource = "GA" Or dataSource = "FB" Then dimensionsArr(dimensionsCount, 2) = LCase$(dimensionsArr(dimensionsCount, 2))

                If dataSource = "GA" Or dataSource = "FB" Then
                    dimensionsArr(dimensionsCount, 3) = .Cells(dimensionNum + dimensionsListStart.row - 1, dimensionsListStart.Column + 2).value   'name incl sub

                    'dimension made of multiple subdimensions
                    If .Cells(dimensionNum + dimensionsListStart.row - 1, dimensionsListStart.Column + 3).value <> "" Then
                        dimensionsCountInclSubGlobal = dimensionsCountInclSubGlobal + .Cells(dimensionNum + dimensionsListStart.row - 1, dimensionsListStart.Column + 3).value
                        dimensionsArr(dimensionsCount, 4) = .Cells(dimensionNum + dimensionsListStart.row - 1, dimensionsListStart.Column + 3).value   'subdimcount
                    Else
                        dimensionsArr(dimensionsCount, 4) = 1
                        dimensionsCountInclSubGlobal = dimensionsCountInclSubGlobal + 1
                    End If




                Else
                    dimensionsArr(dimensionsCount, 3) = dimensionsArr(dimensionsCount, 2)
                    dimensionsArr(dimensionsCount, 4) = 1
                    dimensionsCountInclSubGlobal = dimensionsCountInclSubGlobal + 1
                End If

                If isTime(dimensionsArr(dimensionsCount, 2), , False) Then
                    timeDimensionIncluded = True
                    timeDimensionsIncludedInclMisc = True

                    If mostGranularTimeDimension = "" And isTime(dimensionsArr(dimensionsCount, 2), "year") Then mostGranularTimeDimension = "year"
                    If (mostGranularTimeDimension = "" Or isTime(mostGranularTimeDimension, "year")) And isTime(dimensionsArr(dimensionsCount, 2), "month") Then mostGranularTimeDimension = "month"
                    If (mostGranularTimeDimension = "" Or isTime(mostGranularTimeDimension, "year") Or isTime(mostGranularTimeDimension, "month")) And isTime(dimensionsArr(dimensionsCount, 2), "week") Then mostGranularTimeDimension = "week"
                    If (mostGranularTimeDimension = "" Or isTime(mostGranularTimeDimension, "year") Or isTime(mostGranularTimeDimension, "month") Or isTime(mostGranularTimeDimension, "week")) And isTime(dimensionsArr(dimensionsCount, 2), "date") Then mostGranularTimeDimension = "date"
                    If isTime(dimensionsArr(dimensionsCount, 2), "hour") Then mostGranularTimeDimension = "hour"
                    If isTime(dimensionsArr(dimensionsCount, 2), "date") Then dateDimensionIncluded = True

                    If LCase(dimensionsArr(dimensionsCount, 3)) = "yearmonth" Then
                        If LCase(segmDimName) = "year" Or LCase(segmDimName2) = "year" Then
                            dimensionsArr(dimensionsCount, 2) = "Month"  'if segmdim includes year, then don't include year with month on dim
                            dimensionsArr(dimensionsCount, 3) = "Month"
                        End If
                    ElseIf LCase(dimensionsArr(dimensionsCount, 3)) = "yearweek" Then
                        If dataSource = "GA" Then
                            weekType = "US"
                        Else
                            weekType = "ISO"
                        End If
                        If LCase(segmDimName) = "year" Or LCase(segmDimName2) = "year" Then
                            dimensionsArr(dimensionsCount, 2) = "Week"  'if segmdim includes year, then don't include year with month on dim
                            dimensionsArr(dimensionsCount, 3) = "Week"
                        End If
                    ElseIf LCase(dimensionsArr(dimensionsCount, 3)) = "yearweekiso" Then
                        weekType = "ISO"
                        If LCase(segmDimName) = "year" Or LCase(segmDimName2) = "year" Then
                            dimensionsArr(dimensionsCount, 2) = "Weekiso"  'if segmdim includes year, then don't include year with month on dim
                            dimensionsArr(dimensionsCount, 3) = "Weekiso"
                            If LCase(segmDimName) = "year" Then
                                segmDimName = "YearOfISOweek"
                            Else
                                segmDimName2 = "YearOfISOweek"
                            End If
                        End If
                    End If
                ElseIf isTime(dimensionsArr(dimensionsCount, 2), , True) Then
                    timeDimensionsIncludedInclMisc = True  'day of week/month, quarter
                ElseIf LCase(dimensionsArr(dimensionsCount, 3)) = "nthday" Or LCase(dimensionsArr(dimensionsCount, 3)) = "nthweek" Or LCase(dimensionsArr(dimensionsCount, 3)) = "nthmonth" Then
                    timeDimensionsIncludedInclMisc = True    'will make sorting alphebetic
                Else
                    nonTimeDimensionIncluded = True
                End If


                If dimensions = vbNullString Then
                    dimensions = dimensionsArr(dimensionsCount, 3)
                Else
                    dimensions = dimensions & rscL1 & dimensionsArr(dimensionsCount, 3)
                End If


                tempStr = LCase(dimensionsArr(dimensionsCount, 2))
                If tempStr = "visitlengthcategorized" Then
                    dimensionsRequiringCompressionIncluded = True
                    tempStr2 = "visitlengthcat"
                    columnModificationsStr = columnModificationsStr & "%" & tempStr2 & "->" & dimensionsCountInclSubGlobal & "%"
                End If

            End If

        Next dimensionNum

    End With



    If rawDataReport Then

    End If

    If sendMode = True Then Call checkE(email, dataSource)
    Debug.Print "Dimensions: " & dimensions
    dimensionsBasicStr = dimensions



    stParam1 = "8.09"



    segmDimIncludesDate = False
    segmDimIncludesWeek = False
    segmDimIncludesMonth = False
    segmDimIncludesYear = False


    If queryType = "SD" Then

        If isTime(segmDimName, "", False) Then
            segmDimIsTime = True
        ElseIf isTime(segmDimName2, "", False) Then
            segmDimIsTime = True
        Else
            segmDimIsTime = False
        End If

        If Not isTime(segmDimName, "", False) Then
            segmDimHasNonTimeComponent = True
        ElseIf segmDimCount = 2 And Not isTime(segmDimName2, "", False) Then
            segmDimHasNonTimeComponent = True
        Else
            segmDimHasNonTimeComponent = False
        End If

        If isTime(segmDimName, "year") Then
            segmDimNumForYear = 1
            segmDimIncludesYear = True
            If mostGranularTimeDimension = vbNullString Then mostGranularTimeDimension = "year"
            '  If weekType = "ISO" Then segmDimName = "YearOfISOweek"
        ElseIf isTime(segmDimName2, "year") Then
            segmDimNumForYear = 2
            segmDimIncludesYear = True
            If mostGranularTimeDimension = vbNullString Then mostGranularTimeDimension = "year"
        Else
            segmDimIncludesYear = False
        End If


        If isTime(segmDimName, "month") Then
            segmDimNumForMonth = 1
            segmDimIncludesMonth = True
            If mostGranularTimeDimension = vbNullString Or isTime(mostGranularTimeDimension, "year") Then mostGranularTimeDimension = "month"
        ElseIf isTime(segmDimName2, "month") Then
            segmDimNumForMonth = 2
            segmDimIncludesMonth = True
            If mostGranularTimeDimension = vbNullString Or isTime(mostGranularTimeDimension, "year") Then mostGranularTimeDimension = "month"
        Else
            segmDimIncludesMonth = False
        End If


        If isTime(segmDimName, "week") Then
            segmDimNumForWeek = 1
            segmDimIncludesWeek = True
            If mostGranularTimeDimension = vbNullString Or isTime(mostGranularTimeDimension, "year") Or isTime(mostGranularTimeDimension, "month") Then mostGranularTimeDimension = LCase(segmDimName)
        ElseIf isTime(segmDimName2, "week") Then
            segmDimNumForWeek = 2
            segmDimIncludesWeek = True
            If mostGranularTimeDimension = vbNullString Or isTime(mostGranularTimeDimension, "year") Or isTime(mostGranularTimeDimension, "month") Then mostGranularTimeDimension = LCase(segmDimName2)
        Else
            segmDimIncludesWeek = False
        End If


        If isTime(segmDimName, "date") Then
            segmDimNumForDate = 1
            segmDimIncludesDate = True
            If mostGranularTimeDimension = vbNullString Or isTime(mostGranularTimeDimension, "year") Or isTime(mostGranularTimeDimension, "month") Or isTime(mostGranularTimeDimension, "week") Then mostGranularTimeDimension = "date"
        ElseIf isTime(segmDimName2, "date") Then
            segmDimNumForDate = 2
            segmDimIncludesDate = True
            If mostGranularTimeDimension = vbNullString Or isTime(mostGranularTimeDimension, "year") Or isTime(mostGranularTimeDimension, "month") Or isTime(mostGranularTimeDimension, "week") Then mostGranularTimeDimension = "date"
        Else
            segmDimIncludesDate = False
        End If


        For i = 1 To 2
            If i = 1 Then
                tempStr = LCase(segmDimName)
            Else
                tempStr = LCase(segmDimName2)
            End If
            If tempStr = "visitlengthcategorized" Then
                tempStr2 = "visitlengthcat"
                columnModificationsStr = columnModificationsStr & "%" & tempStr2 & "->" & dimensionsCountInclSubGlobal + i & "%"
                columnModificationsStr = columnModificationsStr & "%" & tempStr2 & "sd->" & i & "%"
                dimensionsRequiringCompressionInSD = True
                dimensionsRequiringCompressionIncluded = True
            End If
        Next i


        If segmDimCount = 2 Then
            segmDimNameCombDisp = Range("segmDimNameDisp").value & " | " & Range("segmDimNameDisp2").value
            segmDimNameComb = segmDimName & rscL1 & segmDimName2
            dimensions = dimensions & rscL1 & segmDimNameComb
        ElseIf fieldNameIsOk(Range("segmDimName").value) = True Then
            segmDimNameCombDisp = Range("segmDimNameDisp").value
            segmDimNameComb = segmDimName
            dimensions = dimensions & rscL1 & segmDimNameComb
        ElseIf fieldNameIsOk(Range("segmDimName2").value) = True Then
            segmDimNameCombDisp = Range("segmDimNameDisp2").value
            segmDimNameComb = segmDimName2
            dimensions = dimensions & rscL1 & segmDimNameComb
        Else
            queryType = "D"
            Debug.Print "There is something strange about this query..."
        End If

        '  segmDimName = segmDimNameComb
        '  segmDimNameDisp = segmDimNameCombDisp
        ReDim segmDimValuesArr(1 To segmDimCount)
    Else
        segmDimIsTime = False
    End If


    If dimensionsRequiringCompressionIncluded = True And uniqueCountMetricsIncluded = True Then
        MsgBox "Illegal combination of dimensions and metrics.", , "Illegal field combination"
        Call hideProgressBox
        End
    End If


    stParam1 = "8.091"

    Dim noDimensions As Boolean
    If dimensionsCount = 0 Then noDimensions = True


    'if segmenting dimensions selected use those as regular dimensions and change query type
    If (rawDataReport Or noDimensions) And queryType = "SD" Then
        If rawDataReport Then
            Debug.Print "Query type marked as SD for raw data report, changing to D"
        Else
            Debug.Print "Query type marked as SD but only segmenting dimensions selected, changing to D"
        End If
        queryType = "D"
        'dimensions = vbNullString
        columnModificationsStr = vbNullString
        folderDimensionIncluded = False
        segmDimIsTime = False


        For segmDimNum = 1 To segmDimCount
            dimensionsCount = dimensionsCount + 1
            dimensionsCountInclSubGlobal = dimensionsCountInclSubGlobal + 1

            dimensionsArr(dimensionsCount, 4) = 1

            Select Case segmDimNum
            Case 1
                dimensionsArr(dimensionsCount, 1) = Range("segmDimNameDisp").value    'name disp
                If dataSource = "GA" Then
                    dimensionsArr(dimensionsCount, 2) = LCase$(Range("segmDimName").value)    'name
                Else
                    dimensionsArr(dimensionsCount, 2) = Range("segmDimName").value    'name
                End If
            Case 2
                dimensionsArr(dimensionsCount, 1) = Range("segmDimNameDisp2").value    'name disp
                If dataSource = "GA" Then
                    dimensionsArr(dimensionsCount, 2) = LCase$(Range("segmDimName2").value)    'name
                Else
                    dimensionsArr(dimensionsCount, 2) = Range("segmDimName2").value    'name
                End If
            End Select

            If isTime(dimensionsArr(dimensionsCount, 2), , False) Then
                timeDimensionIncluded = True
                If mostGranularTimeDimension = "" And LCase(dimensionsArr(dimensionsCount, 2)) = "year" Then mostGranularTimeDimension = "year"
                If (mostGranularTimeDimension = "" Or isTime(mostGranularTimeDimension, "year")) And (LCase(dimensionsArr(dimensionsCount, 2)) = "month" Or LCase(dimensionsArr(dimensionsCount, 2)) = "yearmonth") Then mostGranularTimeDimension = "month"
                If (mostGranularTimeDimension = "" Or isTime(mostGranularTimeDimension, "year") Or isTime(mostGranularTimeDimension, "month")) And (LCase(dimensionsArr(dimensionsCount, 2)) = "week" Or LCase(dimensionsArr(dimensionsCount, 2)) = "weekiso") Then mostGranularTimeDimension = "week"
                If (mostGranularTimeDimension = "" Or isTime(mostGranularTimeDimension, "year") Or isTime(mostGranularTimeDimension, "month") Or isTime(mostGranularTimeDimension, "week") Or mostGranularTimeDimension = "weekiso") And LCase(dimensionsArr(dimensionsCount, 2)) = "date" Then mostGranularTimeDimension = "date"
                If LCase(dimensionsArr(dimensionsCount, 2)) = "hour" Then mostGranularTimeDimension = "hour"
            Else
                nonTimeDimensionIncluded = True
            End If

            dimensionsBasicStr = dimensions

            If debugMode Then Debug.Print "Dimensions after change: " & dimensions

        Next segmDimNum

        If timeDimensionIncluded = True Then timeDimensionsIncludedInclMisc = True

        Call determineHeaderRows

        segmDimCount = 0
        segmDimCategoriesCount = 1

    ElseIf noDimensions And Not rawDataReport Then
        Debug.Print "Query type marked as SD or D but no dimensions selected, changing to A"
        'if no dimensions selected then run aggregate query instead
        Application.StatusBar = False
        Call deleteFromQueryStorageByID(sheetID)
        Call removeSheet
        Range("querytype").value = "A"
        Call aggregateQuery
        End
    ElseIf noDimensions And rawDataReport Then
        'adding placeholder dimension
        dimensionsCount = dimensionsCount + 1
        dimensionsCountInclSubGlobal = dimensionsCountInclSubGlobal + 1

        dimensionsArr(dimensionsCount, 1) = ""
        dimensionsArr(dimensionsCount, 2) = "all"
        dimensionsArr(dimensionsCount, 3) = "all"
        dimensionsArr(dimensionsCount, 4) = 1

        dimensions = "all"
        dimensionsBasicStr = dimensions
    End If



    stParam1 = "8.092"

    For metricSetNum = 1 To metricSetsCount
        metricSetsArr(metricSetNum, 6) = dimensionsBasicStr
    Next metricSetNum


    Dim firstDimensionCountMetric As String

    ReDim dimCountColumnsArr(0)

    dimensionsCombinedCol = resultStartColumn + dimensionsCount + extraDimensionColumns
    firstMetricCol = dimensionsCombinedCol + 1

    If dataSource = "GA" Then
        dimensionCountMetricIncluded = False
        subDimensionCountOrigForLastDim = dimensionsArr(dimensionsCount, 4)
        col = dimensionsCountInclSubGlobal + segmDimCount
        metricSetNum = 1
        For metricNum = 1 To metricsCount
            If metricSetNum < metricsArr(metricNum, 11) Then
                metricSetsArr(metricSetNum, 5) = columnModificationsArr
                Erase columnModificationsArr
                ReDim columnModificationsArr(1 To 30, 1 To 4)
                '1 column
                '2 metric name
                '3 related column
                '4 type: dimCountMetric
                metricSetNum = metricsArr(metricNum, 11)
                metricSetsArr(metricSetNum, 5) = columnModificationsArr
                col = dimensionsCountInclSubGlobal + segmDimCount
            End If
            metricSetNum = metricsArr(metricNum, 11)
            col = col + metricsArr(metricNum, 4)
            If metricsArr(metricNum, 9) <> vbNullString Then
                dimensionCountMetricIncluded = True
                firstDimensionCountMetric = metricsArr(metricNum, 1)
                '   dimensionsArr(dimensionsCount, 4) = dimensionsArr(dimensionsCount, 4) + 1
                col = col + 1
                metricSetsArr(metricSetNum, 6) = metricSetsArr(metricSetNum, 6) & rscL1 & metricsArr(metricNum, 9)  'dimension used

                columnModificationsArr(col, 1) = col - metricsArr(metricNum, 4) + 1    'metric col
                columnModificationsArr(col, 2) = metricsArr(metricNum, 2)   'metric name
                columnModificationsArr(col, 3) = dimensionsCountInclSubGlobal + 1  'dimension col
                columnModificationsArr(col, 4) = "dimCountMetric"  'type
                ReDim Preserve dimCountColumnsArr(UBound(dimCountColumnsArr) + 1)
                dimCountColumnsArr(UBound(dimCountColumnsArr)) = columnModificationsArr(col, 3)

                metricSetsArr(metricSetNum, 5) = columnModificationsArr

            End If
        Next metricNum

        If folderDimensionIncluded = True And dimensionCountMetricIncluded = True Then
            MsgBox "Illegal combination of dimensions and metrics. The directory level dimensions can not be combined in the same report with " & Chr(34) & firstDimensionCountMetric & Chr(34), , "Illegal field combination"
            Call hideProgressBox
            End
        End If

    End If

    '    For rivi = 1 To 30
    '        Debug.Print rivi & " " & columnModificationsArr(rivi, 1) & " " & columnModificationsArr(rivi, 3)
    '    Next rivi


    For metricSetNum = 1 To metricSetsCount
        If segmDimCount = 2 Then
            segmDimNameCombDisp = Range("segmDimNameDisp").value & " | " & Range("segmDimNameDisp2").value
            segmDimNameComb = segmDimName & rscL1 & segmDimName2
            metricSetsArr(metricSetNum, 6) = metricSetsArr(metricSetNum, 6) & rscL1 & segmDimNameComb
        ElseIf segmDimCount = 1 And fieldNameIsOk(segmDimName) = True Then
            segmDimNameCombDisp = Range("segmDimNameDisp").value
            segmDimNameComb = segmDimName
            metricSetsArr(metricSetNum, 6) = metricSetsArr(metricSetNum, 6) & rscL1 & segmDimNameComb
        ElseIf segmDimCount = 1 And fieldNameIsOk(segmDimName2) = True Then
            segmDimNameCombDisp = Range("segmDimNameDisp2").value
            segmDimNameComb = segmDimName2
            metricSetsArr(metricSetNum, 6) = metricSetsArr(metricSetNum, 6) & rscL1 & segmDimNameComb
        End If

        If debugMode = True Then
            Debug.Print "MetricSet " & metricSetNum & " start: " & metricSetsArr(metricSetNum, 1) & " end: " & metricSetsArr(metricSetNum, 2) & " met: " & metricSetsArr(metricSetNum, 4) & " dim: " & metricSetsArr(metricSetNum, 6)
        End If

    Next metricSetNum






    sortType = Range("sortType").value  'custom sort type has been specified in querystorage

    If sortType = vbNullString And updatingPreviouslyCreatedSheet = True Then sortType = fetchValue("sortType", dataSheet)  'take sort type from report sheet

    If updatingPreviouslyCreatedSheet = False Or sortType = vbNullString Then  'set sort type automatically
        If timeDimensionsIncludedInclMisc = True Then
            sortType = "alphabetic"
        Else
            sortType = "metric desc"
        End If
    End If


    stParam1 = "8.10"

    If sendMode = True Then Call checkE(email, dataSource)

    stParam1 = "8.101"

    comparisonType = Range("comparisonType").value
    If comparisonType = "none" Or comparisonType = "" Then
        doComparisons = 0
        iterationsCount = 1
    Else
        If comparisonType = "yearly" Then
            stParam1 = "8.102"
            Range("startDateComparisonCQ").value = DateSerial(Year(startDate1) - 1, Month(startDate1), Day(startDate1))
            Range("endDateComparisonCQ").value = DateSerial(Year(endDate1) - 1, Month(endDate1), Day(endDate1))
            If timeDimensionIncluded = True Then Range("endDateComparisonCQ").value = endDate1 - 350
        ElseIf comparisonType = "previous" Then
            stParam1 = "8.103"
            If timeDimensionIncluded = True Or segmDimIsTime = True Then
                Range("endDateComparisonCQ").value = endDate1
                If isTime(mostGranularTimeDimension, "hour") Then
                    Range("startDateComparisonCQ").value = startDate1
                    Range("endDateComparisonCQ").value = endDate1
                ElseIf isTime(mostGranularTimeDimension, "date") Then
                    Range("startDateComparisonCQ").value = startDate1 - 5
                ElseIf isTime(mostGranularTimeDimension, "week") Then
                    Range("startDateComparisonCQ").value = startDate1 - 15
                ElseIf isTime(mostGranularTimeDimension, "month") Then
                    Range("startDateComparisonCQ").value = startDate1 - 60
                ElseIf isTime(mostGranularTimeDimension, "year") Then
                    Range("startDateComparisonCQ").value = startDate1 - 400
                Else
                    Range("endDateComparisonCQ").value = startDate1 - 1
                    Range("startDateComparisonCQ").value = Range("endDateComparisonCQ").value - (endDate1 - startDate1)
                End If
            Else
                Range("endDateComparisonCQ").value = startDate1 - 1
                Range("startDateComparisonCQ").value = Range("endDateComparisonCQ").value - (endDate1 - startDate1)
            End If
        End If


        stParam1 = "8.104"

        doComparisons = 1
        startDate2 = Range("startDateComparisonCQ").value
        endDate2 = Range("endDateComparisonCQ").value
        iterationsCount = 2



        If dataSource = "GA" Then
            If startDate2 < DateSerial(2005, 1, 1) Then startDate2 = DateSerial(2005, 1, 1)
            If endDate2 < DateSerial(2005, 1, 1) Then endDate2 = DateSerial(2005, 1, 1)
        End If



        If startDate2 > endDate2 Then
            MsgBox "Invalid comparison date range (start date should be before end date)"
            Call hideProgressBox
            End
        End If
    End If

    stParam1 = "8.105"

    Debug.Print "Date range: " & startDate1 & "-" & endDate1
    If comparisonType <> "none" Then Debug.Print "Comparison date range: "; startDate2 & "-" & endDate2 & " (comparison type " & comparisonType & ")"
    If timeDimensionIncluded = True Then Debug.Print "Most granular time dimension: " & mostGranularTimeDimension


    stParam1 = "8.106"





    Exit Sub


generalErrHandler:

    stParam2 = "REPORTINITERROR " & Err.Number & "|" & Err.Description & "|" & Application.StatusBar
    Debug.Print "REPORTINITERROR: " & stParam1 & " " & stParam2

    'Call checkE(email, dataSource, True)


    If Err.Number = 18 Then
        Call hideProgressBox
        Call removeTempsheet
        End
    End If


    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
    Resume Next
    'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


End Sub


