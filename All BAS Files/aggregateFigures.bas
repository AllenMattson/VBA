Attribute VB_Name = "aggregateFigures"
Option Private Module
Option Explicit


Sub fetchAggregateFigures()
    Call fetchAggregateFiguresInitAndFetch
    Call fetchAggregateFiguresFormatting
End Sub

Sub fetchAggregateFiguresInitAndFetch()

    Dim buttonObjPrev As Object

    Dim doConditionalFormatting As Boolean

    Dim dimensionCountMetricIncludedInMetricSet As Boolean

    Dim tempArr As Variant
    showNoteStr = ""


    dimensionCountMetricIncludedInMetricSet = False
    givemaxResultsPerQueryWarning = False
    stParam1 = "7.00"
    stParam4 = vbNullString

    Application.EnableEvents = False

    Application.ScreenUpdating = False


    aika1 = Timer


    If Range("sumAllProfiles").value = True Then
        sumAllProfiles = True
    Else
        sumAllProfiles = False
    End If
    If sumAllProfiles Then
        allProfilesInOneQuery = True
    Else
        allProfilesInOneQuery = False
    End If

    advancedSettings = Range("advancedSettings").value


    includeOther = False



    Randomize
    Call checkOperatingSystem

    If useQTforDataFetch And Not usingMacOSX And Int((5 * Rnd) + 1) = 1 Then
        Call testConnection  'every fifth run, test if MSXML available
    End If

    On Error GoTo generalErrHandler
    If debugMode = True Then On Error GoTo 0
    Application.EnableCancelKey = xlErrorHandler

    Debug.Print ""
    Debug.Print "------------------------------------------------------"
    Debug.Print "------------NEW QUERY A " & Now & "------------"
    Debug.Print "------------------------------------------------------"


    queryType = "A"

    reportContainsSampledData = False
    dimensionsRequiringCompressionInSD = False

    '  numberOfCharsThatCanBeReturnedToCell = testNumberOfCharsThatCanBeReturnedToCell()

    Call getProxySettingsIfNeeded


    dataSource = Range("dataSource").value
    clientLoginModeForGA = Range("clientLoginModeForGA").value

    Call setDatasourceVariables

    fontName = Range("mainFont").value


    dateRangeType = Range("daterangetype").value
    If dateRangeType = vbNullString Or dateRangeType = "custom" Then dateRangeType = "fixed"

    If dateRangeType = "fixed" Then
        Dim startDateRange As Range
        Set startDateRange = Range("startDate" & varsuffix)

        Dim endDateRange As Range
        Set endDateRange = Range("endDate" & varsuffix)

        If startDateRange.value = vbNullString Or endDateRange.value = vbNullString Then
            MsgBox "Choose dates first"
            Call hideProgressBox
            End
        End If

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





    If runningSheetRefresh = True Then
        If Range("loggedin" & varsuffix).value <> True Then
            Call hideProgressBox
            If importingFromOldVersion = True Then
                MsgBox "Unable to copy report as you have not logged in with " & serviceName & ". Log in and try again."
            Else
                MsgBox "Unable to refresh report as you have not logged in with " & serviceName & ". Log in and try again."
            End If
            End
        End If
    End If



    '    If dataSource = "GW" Then
    '        If Range("loggedin").value = False Then
    '            MsgBox "You need to be logged in to the Google Analytics Module in order to run Webmaster Tools reports"
    '            End
    '        End If
    '    End If


    Application.ScreenUpdating = False


    Range("queryRunTime").value = Now()

    doHyperlinks = Range("doHyperlinks").value

    If usingMacOSX = False Then
        ProgressBox.Show False
        ProgressBox.tweetLink.Visible = True
    End If
    progresspct = 3
    Call updateProgress(progresspct, "Authenticating to " & serviceName & "...")

    Dim dataRivi As Long

    Dim dataSar As Long

    Dim rivi As Long
    Dim sar As Long
    Dim arrRivi As Long
    Dim i As Long
    Dim j As Long
    Dim profName As String
    Dim accountName As String
    Dim num As Variant

    processQueriesTotal = 0
    processQueriesCompleted = 0
    processIDsStr = ""
    objHTTPstatusRunning = False
    nameEncodingStr = ""


    metricsCount = 0

    Dim metricsListStart As Range
    Set metricsListStart = Range("metric1name" & varsuffix)
    Dim metricsListStartColumn As Long
    Dim metricsListStartRow As Long
    metricsListStartColumn = metricsListStart.Column
    metricsListStartRow = metricsListStart.row

    Dim metricNumResultArr As Long

    Dim tempStr As String
    Dim tempStr2 As String


    Dim arvo As Variant
    Dim arvo1 As Variant    'value from first iteration when comparison is used
    Dim div As Variant

    Dim col As Long


    folderDimensionIncluded = False


    If dataSource = "GA" Then
        If Range("segmentID").value = vbNullString Or Range("segmentID").value = 0 Then
            segmentCount = 1
            ReDim segmentArr(1 To segmentCount, 1 To 2)
            segmentArr(1, 1) = "-1"
            segmentArr(1, 2) = "All Visits"
            segmentIsAllVisits = True
        Else
            tempStr = Replace(Range("segmentName").value, ", ", ",")
            segmentID = Range("segmentID").value
            segmentCount = UBound(Split(Range("segmentID").value, ",")) + 1
            ReDim segmentArr(1 To segmentCount, 1 To 2)
            For segmentNum = 1 To segmentCount
                segmentArr(segmentNum, 1) = Split(segmentID, ",")(segmentNum - 1)
                segmentArr(segmentNum, 2) = Split(tempStr, ",")(segmentNum - 1)
            Next segmentNum
            segmentIsAllVisits = False
        End If
    Else
        segmentCount = 1
        ReDim segmentArr(1 To segmentCount, 1 To 2)
        segmentArr(1, 1) = "-1"
        segmentArr(1, 2) = "All Visits"
        segmentIsAllVisits = True
    End If


    If sumAllProfiles And segmentCount = 1 Then
        doTotals = False
    Else
        doTotals = Range("doTotals").value
    End If



    sortType = Range("sortType").value  'custom sort type has been specified in querystorage

    If sortType = vbNullString And updatingPreviouslyCreatedSheet = True Then sortType = fetchValue("sortType", dataSheet)  'take sort type from report sheet

    If updatingPreviouslyCreatedSheet = False Or sortType = vbNullString Then  'set sort type automatically
        sortType = "metric desc"
    End If



    Set profileListStart = Range("profileListStart" & varsuffix)

    comparisonValueType = Range("comparisonValueType").value
    If comparisonValueType = vbNullString Then comparisonValueType = "perc"
    createCharts = Range("createCharts").value



    If Range("separeteQueryForEachGAMetric").value = True Then
        separeteQueryForEachGAMetric = True
    Else
        separeteQueryForEachGAMetric = False
    End If



    If Range("conditionalFormattingType").value <> "none" Then doConditionalFormatting = True


    ReDim columnModificationsArr(1 To 30, 1 To 4)
    '1 column
    '2 metric name
    '3 target column


    dimensionsCountInclSubGlobal = 0
    dimensionsCount = 0
    segmDimCount = 0
    dimensionCountMetricIncluded = False

    filterStr = Range("filterstring").value





    comparisonType = Range("comparisonType").value


    If comparisonType = "none" Then
        doComparisons = 0
        iterationsCount = 1
    Else
        If comparisonType = "yearly" Then
            Range("startDateComparisonCQ").value = DateSerial(Year(startDate1) - 1, Month(startDate1), Day(startDate1))
            Range("endDateComparisonCQ").value = DateSerial(Year(endDate1) - 1, Month(endDate1), Day(endDate1))
        ElseIf comparisonType = "previous" Then
            Range("endDateComparisonCQ").value = startDate1 - 1
            Range("startDateComparisonCQ").value = Range("endDateComparisonCQ").value - (endDate1 - startDate1)
        End If

        doComparisons = 1
        startDate2 = Range("startDateComparisonCQ").value
        endDate2 = Range("endDateComparisonCQ").value
        iterationsCount = 2

        If startDate2 > endDate2 Then
            MsgBox "Invalid comparison date range (start date should be before end date)"
            Call hideProgressBox
            End
        End If
    End If




    givemaxResultsPerQueryWarning = False
    giveUniqueSumWarning = False

    maxResults1 = Range("maxresults").value
    maxResults = maxResults1

    maxResultsMultiplierForComparisonQuery = 2

    If maxResults1 <= 10000 Then
        maxResultsMultiplierForComparisonQuery = 2
    Else
        maxResultsMultiplierForComparisonQuery = 1
    End If

    avoidSampling = Range("avoidSampling").value


    Dim muutos As Variant

    Dim sheetName As String
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
            If dataSource = "GA" Then
                MsgBox "No profiles have been selected. Select at least one from the list and try again."
            ElseIf dataSource = "AW" Or dataSource = "AC" Then
                MsgBox "No accounts have been selected. Select at least one from the list and try again."
            ElseIf dataSource = "FB" Then
                MsgBox "No pages or applications have been selected. Select at least one from the list and try again."
            ElseIf dataSource = "YT" Then
                MsgBox "No videos or channels have been selected. Select at least one from the list and try again."
            End If
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

        If InStr(1, tempStr, "|" & email) = 0 And Int(emailLastCheckedOK) <> Date Or rCount = 0 Then
            Call checkE(email, dataSource, , , True)
            Call storeEmailCheckedDateToSheet(email)
            tempStr = tempStr & "|" & email
            rCount = rCount + 1
            Range("rcount").value = Range("rcount").value + 1
            i = i + 1
        End If

        profilesArr(profNum, 8) = profilesArr(profNum, 3)     'id
        profilesArr(profNum, 8) = profilesArr(profNum, 8) & rscL2 & profilesArr(profNum, 2)      'name
        profilesArr(profNum, 8) = profilesArr(profNum, 8) & rscL2    ' & profilesArr(profNum, 7)      'token
        profilesArr(profNum, 8) = profilesArr(profNum, 8) & rscL2 & profilesArr(profNum, 1)      'account name

    Next profNum
    Erase tempArr
    tempStr = vbNullString


    If allProfilesInOneQuery Then
        allProfilesStr = vbNullString
        For profNum = 1 To profileCount
            If allProfilesStr <> vbNullString Then allProfilesStr = allProfilesStr & rscL1
            allProfilesStr = allProfilesStr & profilesArr(profNum, 3)  'id
            allProfilesStr = allProfilesStr & rscL2 & convertRSCL(profilesArr(profNum, 2))  'name
            allProfilesStr = allProfilesStr & rscL2 & convertRSCL(profilesArr(profNum, 7))  'token
        Next profNum
        allProfilesStr = allProfilesStr
        profileCount = 1
        '        If sumAllProfiles Then
        '            profilesArr(1, 1) = vbNullString  'acc name
        '            profilesArr(1, 2) = vbNullString  'prof name
        '        End If
    End If



    progresspct = 5
    Call updateProgress(progresspct, "Checking selected metrics...")

    goalsIncluded = False
    metricSetNum = 1
    metricSetsCount = 1
    metricNumInMetricSet = 0
    metricNumInclSubInMetricSet = 0


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
                If dataSource = "GA" Or dataSource = "FB" Or dataSource = "GW" Or dataSource = "ST" Then
                    metricsArr(metricsCount, 8) = .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 7).value    'invert condform
                Else
                    metricsArr(metricsCount, 8) = .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 9).value    'invert condform
                End If
                If metricsArr(metricsCount, 8) = vbNullString Then metricsArr(metricsCount, 8) = False
                stParam1 = "8.063"

                If dataSource = "AW" Then
                    metricsArr(metricsCount, 12) = .Cells(metricNum + metricsListStartRow - 1, metricsListStartColumn + 7).value  'don't calculate totals
                End If


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

                stParam1 = "8.064"
                If metricsCount = 1 Then
                    metrics = metricsArr(metricsCount, 3)
                Else
                    metrics = metrics & rscL1 & metricsArr(metricsCount, 3)
                End If


                stParam1 = "8.065"
                Debug.Print metrics
            End If

        Next metricNum

    End With
    stParam1 = "8.066"
    Debug.Print "Metrics: " & metrics

    If metricsCount = 0 Then
        Application.StatusBar = False
        MsgBox "Choose at least one metric first"
        Call hideProgressBox
        End
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
                If metricSetNum > 1 Then arvo = arvo & rscL1 & firstMetricStr   'add first metric to other metric sets to get sorting right
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
        If metricSetNum > 1 Then arvo = arvo & rscL1 & firstMetricStr   'add first metric to other metric sets to get sorting right
        metricSetsArr(metricSetsCount, 4) = arvo
        'm = 2: print m & ": " : print "first: "  & metricSetsArr(m,1) : print "last: "  & metricSetsArr(m,2) : print "preceding count: "  & metricSetsArr(m,3)  : print "metrics: "  & metricSetsArr(m,4)
        If debugMode = True Then
            For metricSetNum = 1 To metricSetsCount
                Debug.Print "MetricSetNum " & metricSetNum & ": " & metricSetsArr(metricSetNum, 4)
            Next metricSetNum
        End If
    Else
        ReDim metricSetsArr(1 To 1, 1 To 10)
        metricSetsCount = 1
        metricSetsArr(1, 1) = 1
        metricSetsArr(1, 2) = metricsCount
        metricSetsArr(1, 3) = 0
        metricSetsArr(1, 4) = metrics
        metricSetsArr(1, 6) = vbNullString
        metricSetsArr(1, 7) = vbNullString
    End If


    Dim firstDimensionCountMetric As String

    ReDim dimCountColumnsArr(0)

    dimensionsCombinedCol = resultStartColumn + dimensionsCount + extraDimensionColumns

    If dataSource = "GA" Then
        dimensionCountMetricIncluded = False
        col = 0
        metricSetNum = 1
        For metricNum = 1 To metricsCount
            If metricSetNum < metricsArr(metricNum, 11) Then
                metricSetsArr(metricSetNum, 5) = columnModificationsArr
                Erase columnModificationsArr
                ReDim columnModificationsArr(1 To 30, 1 To 4)
                '1 column
                '2 metric name
                '3 related column
                '4 type
                col = 0
                metricSetNum = metricsArr(metricNum, 11)
                metricSetsArr(metricSetNum, 5) = columnModificationsArr
            End If
            metricSetNum = metricsArr(metricNum, 11)
            col = col + metricsArr(metricNum, 4)
            If metricsArr(metricNum, 9) <> vbNullString Then
                dimensionCountMetricIncluded = True
                firstDimensionCountMetric = metricsArr(metricNum, 1)
                '  dimensionsArr(dimensionsCount, 4) = dimensionsArr(dimensionsCount, 4) + 1
                col = col + 1
                metricSetsArr(metricSetNum, 6) = metricsArr(metricNum, 9)  'dimension used

                columnModificationsArr(col, 1) = col - metricsArr(metricNum, 4) + 1    'metric col
                columnModificationsArr(col, 2) = metricsArr(metricNum, 2)
                columnModificationsArr(col, 3) = 1  'dimension col
                columnModificationsArr(col, 4) = "dimCountMetric"  'type

                ReDim Preserve dimCountColumnsArr(UBound(dimCountColumnsArr) + 1)
                dimCountColumnsArr(UBound(dimCountColumnsArr)) = columnModificationsArr(col, 3)

                metricSetsArr(metricSetNum, 5) = columnModificationsArr

            End If
        Next metricNum

    End If






    If goalsIncluded = True Then goalsArr = Range("goals").value



    dimensions = ""


    Application.Calculation = xlManual

    If SheetExists(sheetName) = False Then
        Set dataSheet = ThisWorkbook.Sheets.Add
        dataSheet.Name = sheetName
        dataSheet.Tab.ColorIndex = 13
        If Twitter.Visible = xlSheetVisible Then
            dataSheet.move after:=Twitter
        ElseIf TwitterAds.Visible = xlSheetVisible Then
            dataSheet.move after:=TwitterAds
        ElseIf Stripe.Visible = xlSheetVisible Then
            dataSheet.move after:=Stripe
        ElseIf MailChimp.Visible = xlSheetVisible Then
            dataSheet.move after:=MailChimp
        ElseIf Webmaster.Visible = xlSheetVisible Then
            dataSheet.move after:=Webmaster
        ElseIf YouTube.Visible = xlSheetVisible Then
            dataSheet.move after:=YouTube
        ElseIf Facebook.Visible = xlSheetVisible Then
            dataSheet.move after:=Facebook
        ElseIf BingAds.Visible = xlSheetVisible Then
            dataSheet.move after:=BingAds
        ElseIf FacebookAds.Visible = xlSheetVisible Then
            dataSheet.move after:=FacebookAds
        ElseIf AdWords.Visible = xlSheetVisible Then
            dataSheet.move after:=AdWords
        Else
            dataSheet.move after:=Analytics
        End If
        updatingPreviouslyCreatedSheet = False
    Else
        Set dataSheet = ThisWorkbook.Worksheets(sheetName)
        updatingPreviouslyCreatedSheet = True
    End If


    Randomize


    If updatingPreviouslyCreatedSheet = False Then
        sheetID = Range("sheetID").value
    Else
        sheetID = dataSheet.Cells(1, 1).value
        sheetID = findRangeName(dataSheet.Cells(1, 1))
    End If


    Dim resultStart As Range
    If createCharts = True Then
        Set resultStart = dataSheet.Cells(6, reportStartColumn + 15)
    Else
        Set resultStart = dataSheet.Cells(6, reportStartColumn + 10)
    End If



    dataRivi = resultStart.row

    resultStartRow = resultStart.row
    resultStartColumn = resultStart.Column

    firstMetricCol = resultStartColumn + 3
    If segmentCount > 1 Then firstMetricCol = firstMetricCol + 1

    Set tempSheet = ThisWorkbook.Worksheets.Add
    tempSheet.Name = "temp_" & Round(1000000 * Rnd(), 0)
    dataSheet.Select


    Call updateProgress(5, "Initalizing queries...")

    SDlabelsQuery = False
    queriesCompletedCount = 0
    arrRivi = 1
    queryCount = Evaluate(profileCount & "*" & iterationsCount & "*" & metricSetsCount & "*" & segmentCount)


    Call initializeObjHTTP

    ReDim queryArr(1 To queryCount, 1 To 23)
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
    '15 additional query parameters
    '16 error count
    '17 querynum of parent query where subquerynum = 1  (foundDimValuesArr stored there)
    '18 username
    '19 metricSetNum
    '20 queryIDforDB
    '21 segment ID
    '22 SDothersQuery
    '23 contains dimcount metric

    ReDim objHTTParr(1 To maxSimultaneousQueries, 1 To 3)
    '1 in use?
    '2 querynum


    Call initializeFetchArrays

    initialFetchRound = True
    Call runQueriesOnFreeObjHTTPs
    initialFetchRound = False


    progresspct = 6
    Call updateProgress(progresspct, "Marking data headers...")


    With dataSheet

        .Select
        If .FilterMode Then .ShowAllData

        If reportStartColumn > 1 Then
            .Range(.Cells(1, 1), .Cells(1, reportStartColumn - 1)).EntireColumn.Hidden = True
            If usingMacOSX = True Then Call hideProgressBox
        End If


        If updatingPreviouslyCreatedSheet = True Then
            sar = 0
            sar = fetchValue("lastCol", dataSheet)
            If IsNumeric(sar) And sar > 0 Then
                With .Columns(ColumnLetter(resultStartColumn) & ":" & ColumnLetter(sar))
                    .Hidden = False
                    .UnMerge
                    .ClearContents
                    If doConditionalFormatting = True Then .FormatConditions.Delete
                End With
            Else
                With .Columns(ColumnLetter(resultStartColumn) & ":" & ColumnLetter(.Range("A1").SpecialCells(xlCellTypeLastCell).Column))
                    .UnMerge
                    .ClearContents
                    If doConditionalFormatting = True Then .FormatConditions.Delete
                End With
            End If

            Range(sheetID & "_dataRange").Copy tempSheet.Range(Range(sheetID & "_dataRange").Address)

        Else
            .Cells.NumberFormat = ""
        End If

        If updatingPreviouslyCreatedSheet = False Then


            .Cells(1, 1).value = sheetID
            .Cells(1, 1).Name = sheetID
            Call storeValue("sheetID", sheetID, dataSheet)



            Dim buttonNum As Long
            Dim firstButtonLeft As Long



            buttonNum = 2

            firstButtonLeft = Round(.Cells(1, reportStartColumn + 4).Left + buttonSpaceBetween)

            Call updateProgress(progresspct, "Inserting remove sheet button...")

            Set buttonObj = dataSheet.Shapes.AddTextbox(1, 342, 15, 118, 29)
            With buttonObj
                .OnAction = "removeSheet"
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
                .TextFrame.Characters.Text = "REMOVE SHEET"
                .TextFrame.Characters.Font.ColorIndex = 1
                .TextFrame.Characters.Font.Size = 9
                .Fill.ForeColor.RGB = buttonColourRed
                .Line.ForeColor.RGB = buttonBorderColour
                .Height = buttonHeight
                .Width = buttonWidth
                .Top = buttonTop
                .Left = firstButtonLeft + (buttonNum - 1) * (buttonWidth + buttonSpaceBetween)
                .Name = sheetID & "RemoveSheetButton"
            End With
            Call updateProgress(progresspct, "Inserting remove sheet button...")


            If excelVersion <= 11 Then
                .Cells.Interior.ColorIndex = 2
            Else
                .Cells.Interior.Color = Range("sheetBackgroundColour").Interior.Color
            End If

            .Rows(1).RowHeight = 5

        End If


        vsarData = resultStartColumn + 2 + metricsCount + doComparisons * metricsCount
        If segmentCount > 1 Then vsarData = vsarData + 1
        vriviData = resultStartRow + profileCount * segmentCount - 1




        ReDim columnInfoArr(1 To vsarData + 10, 1 To 14)
        '1 metric name
        '2 invert conditional formatting
        '3 profile name
        '4 segmdimname
        '5 hidden
        '6 comparisons
        '7 SD query Other category
        '8 metric code
        '9 metric submetric count
        '10 data fetch error
        '11 metricnum
        '12 prof id
        '13 profnum|metricnum|segmentNum|segmdimcategorynum|iterationnum
        '14 segmentnum

        Call fillDataColumnNumbers



        If updatingPreviouslyCreatedSheet = False Then
            With .Cells(2, reportStartColumn + 1)
                .value = UCase(serviceName & " report")
                ' .Font.Name = fontName
                With .Resize(1, 3)
                    .Interior.ColorIndex = 37
                    .Font.ColorIndex = 2
                End With
            End With
        End If



        .Cells(3, reportStartColumn + 1).value = "Fetched"
        .Cells(3, reportStartColumn + 2).value = Now()
        .Cells(3, reportStartColumn + 3).value = Now()

        .Cells(3, reportStartColumn + 2).NumberFormatLocal = Range("numformatDate").NumberFormatLocal
        .Cells(3, reportStartColumn + 3).NumberFormatLocal = Range("numformatTime").NumberFormatLocal

        stParam1 = "8.1211"
        If dateRangeType = "fixed" Or dateRangeType = "custom" Then
            .Cells(4, reportStartColumn + 1).value = "Date range"
            .Cells(4, reportStartColumn + 1).Font.Bold = True
            .Cells(4, reportStartColumn + 2).value = startDate1
            .Cells(4, reportStartColumn + 3).value = endDate1

            .Cells(4, reportStartColumn + 2).NumberFormatLocal = Range("numformatDate").NumberFormatLocal
            .Cells(4, reportStartColumn + 3).NumberFormatLocal = Range("numformatDate").NumberFormatLocal

            .Cells(4, reportStartColumn + 2).Name = sheetID & "_" & "sdate"
            .Cells(4, reportStartColumn + 3).Name = sheetID & "_" & "edate"

            .Cells(4, reportStartColumn + 2).Interior.ColorIndex = 16
            .Cells(4, reportStartColumn + 3).Interior.ColorIndex = 16

            .Cells(4, reportStartColumn + 2).Font.ColorIndex = 2
            .Cells(4, reportStartColumn + 3).Font.ColorIndex = 2

        Else

            dateRangeTypeDisp = getDispNameForDateRangeType(dateRangeType)
            .Cells(4, reportStartColumn + 1).value = "Report covers " & LCase(dateRangeTypeDisp)
            '  .Cells(4, 3).value = dateRangeTypeDisp
            .Cells(4, reportStartColumn + 2).Font.Bold = True

            .Cells(5, reportStartColumn + 1).value = "Dates"

            .Cells(5, reportStartColumn + 2).value = startDate1
            .Cells(5, reportStartColumn + 3).value = endDate1

            .Cells(5, reportStartColumn + 2).Font.Bold = False
            .Cells(5, reportStartColumn + 3).Font.Bold = False

            .Cells(5, reportStartColumn + 2).NumberFormatLocal = Range("numformatDate").NumberFormatLocal
            .Cells(5, reportStartColumn + 3).NumberFormatLocal = Range("numformatDate").NumberFormatLocal

            .Cells(5, reportStartColumn + 2).Name = sheetID & "_" & "sdate"
            .Cells(5, reportStartColumn + 3).Name = sheetID & "_" & "edate"

        End If



        If updatingPreviouslyCreatedSheet = False Then
            Call storeValue("rowLabelsCol", resultStartColumn + 2, dataSheet)

            Call storeValue("firstCol", resultStartColumn, dataSheet)
            Call storeValue("lastCol", vsarData + 1, dataSheet)

            Call storeValue("sortingCol", 1, dataSheet)
            Call storeValue("sortType", 1, dataSheet)


            .Range(.Cells(1, reportStartColumn), .Cells(10, reportStartColumn + 4)).Font.Size = 9
            .Range(.Cells(2, reportStartColumn + 4), .Cells(3, reportStartColumn + 7)).Font.ColorIndex = 2
        End If

        If updatingPreviouslyCreatedSheet = True Then
            If dateRangeType = "fixed" Or dateRangeType = "custom" Then
                .Cells(5, reportStartColumn + 1).Resize(5, 1).ClearContents
            Else
                .Cells(6, reportStartColumn + 1).Resize(4, 1).ClearContents
            End If
        End If



        If doComparisons = 1 Then
            With .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1)
                If comparisonType = "previous" Then
                    .value = "Changes calculated vs. previous period of same length (" & startDate2 & "-" & endDate2 & ")"
                ElseIf comparisonType = "yearly" Then
                    .value = "Changes calculated vs. same period a year earlier (" & startDate2 & "-" & endDate2 & ")"
                Else
                    .value = "Changes calculated vs. " & startDate2 & "-" & endDate2
                End If
            End With
        End If


        If segmentIsAllVisits = False And segmentCount = 1 Then .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1).value = "Segment: " & Range("segmentname").value
        If filterStr <> vbNullString Then .Cells(vikarivi(.Cells(1, reportStartColumn + 1)) + 1, reportStartColumn + 1).value = "Filter: " & filterStr

        If Not sumAllProfiles Then
            If dataSource = "GA" Then
                .Cells(resultStartRow - 1, resultStartColumn).value = "Profile ID"
                .Cells(resultStartRow - 1, resultStartColumn + 1).value = "Account"
                .Cells(resultStartRow - 1, resultStartColumn + 2).value = "Profile"
            ElseIf dataSource = "AW" Then
                .Cells(resultStartRow - 1, resultStartColumn).value = "Account ID"
                .Cells(resultStartRow - 1, resultStartColumn + 1).value = "MCC"
                .Cells(resultStartRow - 1, resultStartColumn + 2).value = "Account"
            ElseIf dataSource = "AC" Then
                .Cells(resultStartRow - 1, resultStartColumn).value = "Account ID"
                .Cells(resultStartRow - 1, resultStartColumn + 1).value = "Account"
                .Cells(resultStartRow - 1, resultStartColumn + 2).value = "Sub-account"
            End If
        End If

        If segmentCount > 1 Then
            .Cells(resultStartRow - 1, resultStartColumn + 3).value = "Segment"
        End If

        For metricNum = 1 To metricsCount

            With .Cells(resultStartRow - 1, firstMetricCol + metricNum - 1 + doComparisons * (metricNum - 1))
                .value = metricsArr(metricNum, 1)
                If dataSource = "GA" And goalsIncluded = True And metricsArr(metricNum, 10) <> "" And profileCount = 1 Then
                    'goal names
                    profID = profilesArr(1, 3)
                    arvo = getGoalName(profID, metricsArr(metricNum, 10))
                    If arvo <> vbNullString Then .value = metricsArr(metricNum, 1) & ": " & arvo
                    metricsArr(metricNum, 1) = metricsArr(metricNum, 1) & ": " & arvo
                End If
            End With

            If updatingPreviouslyCreatedSheet = False Then
                With .Cells(1, firstMetricCol + (metricNum - 1) + doComparisons * (metricNum - 1)).EntireColumn
                    Select Case metricsArr(metricNum, 6)
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


            If doComparisons = 1 Then
                .Cells(resultStartRow - 1, firstMetricCol + (metricNum - 1) + doComparisons * (metricNum - 1) + 1).value = "*"

                With .Columns(ColumnLetter(firstMetricCol + (metricNum - 1) + doComparisons * (metricNum - 1) + 1))
                    Select Case comparisonValueType
                    Case "perc"
                        .NumberFormat = "0.0 %"  'Range("numFormatChange").NumberFormat
                    Case Else
                        .NumberFormat = .Cells(1, 1).Offset(, -1).NumberFormat
                    End Select
                    .Font.Size = 9
                    .ColumnWidth = 7
                End With
            End If


        Next metricNum


        .Rows(resultStartRow - 1).Font.Bold = True

        If runningSheetRefresh = False Then
            Application.ScreenUpdating = True
            Application.ScreenUpdating = False
        End If







        stParam1 = "7.50"



        Call updateProgress(progresspct, "Fetching data...")





        If debugMode = False Then Application.Cursor = xlWait

        i = 0
        processStatusTimer = Timer
        inDataFetchLoop = True
        Do


            stParam1 = "7.51"

            DoEvents

            If Timer - processStatusTimer >= 2 Or objHTTPstatusRunning Then
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


            DoEvents
            DoEvents

            stParam1 = "7.52"
            'checks for completed queries, stores results and frees up objhttp
            If allQueriesFetched = False Then Call checkForCompletedObjHTTPs
            stParam1 = "7.53"
            'checks for free objhttps and runs queries
            If allQueriesStarted = False Then Call runQueriesOnFreeObjHTTPs
            stParam1 = "7.54"



            'parse results
            For queryNum = 1 To queryCount
                If queryArr(queryNum, 9) = True And queryArr(queryNum, 11) = False Then
                    Debug.Print "Parsing response for query " & queryNum
                    Call updateProgressAdditionalMessage("Parsing response data")
                    If IsArray(arr) Then Erase arr

                    profNum = queryArr(queryNum, 1)
                    iterationNum = queryArr(queryNum, 4)
                    SDlabelsQuery = queryArr(queryNum, 3)



                    '
                    '                    If dataSource = "FB" Then
                    '                        stParam1 = "8.162341"
                    '                        arr = parseFBResponse(queryArr(queryNum, 8))
                    '                    Else
                    stParam1 = "7.541"
                    arr = parseResponse(queryArr(queryNum, 8))
                    '       End If

                    queryArr(queryNum, 8) = arr
                    queryArr(queryNum, 11) = True

                    Call checkArrForErrors
                End If
            Next queryNum

            stParam1 = "7.57"
            DoEvents


            'checks queries in order, if finished then places data into sheet, exits loop when first unfinished query found
            foundNonFinishedQuery = False
            For queryNum = 1 To queryCount

                If queryArr(queryNum, 11) = True Then
                    If queryArr(queryNum, 10) = False Then
                        iterationNum = queryArr(queryNum, 4)


                        If foundNonFinishedQuery = False Or iterationNum = 1 Then
                            stParam1 = "7.58"

                            queryArr(queryNum, 10) = True
                            queriesCompletedCount = queriesCompletedCount + 1

                            Debug.Print "Started processing to sheet: " & queryNum
                            Call updateProgressAdditionalMessage("Processing data into sheet")
                            If IsArray(arr) Then Erase arr
                            arr = queryArr(queryNum, 8)
                            queryArr(queryNum, 8) = ""

                            profNum = queryArr(queryNum, 1)

                            metricSetNum = queryArr(queryNum, 19)

                            SDlabelsQuery = queryArr(queryNum, 3)


                            profID = profilesArr(profNum, 3)
                            profName = profilesArr(profNum, 2)
                            accountName = profilesArr(profNum, 1)
                            email = profilesArr(profNum, 4)

                            segmentNum = queryArr(queryNum, 21)
                            segmentName = segmentArr(segmentNum, 2)

                            stParam1 = "7.59"


                            If processQueriesTotal > 0 Then
                                progresspct = 10 + 20 * queriesCompletedCount / queryCount + 50 * processQueriesCompleted / processQueriesTotal
                            Else
                                progresspct = 10 + 20 * queriesCompletedCount / queryCount
                            End If
                            Call updateProgress(progresspct, "Fetching & processing data...", "Processing data into sheet")



                            If iterationNum = 1 Then
                                startDate = startDate1
                                endDate = endDate1
                                maxResults = 1
                            Else
                                startDate = startDate2
                                endDate = endDate2
                                maxResults = 1
                            End If

                            dataRivi = resultStartRow + (profNum - 1) * segmentCount + (segmentNum - 1)




                            stParam1 = "7.60"

                            If iterationNum = 1 Then
                                If Not sumAllProfiles Then
                                    If dataSource <> "GW" Then .Cells(dataRivi, resultStartColumn).value = profID    ' CStr(Int((9999999 - 100000 + 1) * Rnd + 100000))
                                    If dataSource <> "GW" Then .Cells(dataRivi, resultStartColumn + 1).value = Left$(accountName, 255)
                                    .Cells(dataRivi, resultStartColumn + 2).value = Left$(profName, 255)

                                    If doHyperlinks Then
                                        If dataSource = "YT" Then
                                            .Hyperlinks.Add Cells(dataRivi, resultStartColumn + 1), "http://www.youtube.com/user/" & accountName    ', "Open channel in browser"
                                            If profID <> "TOTALS" Then
                                                .Hyperlinks.Add Cells(dataRivi, resultStartColumn + 2), "http://www.youtube.com/watch?v=" & profID    ', "Open video in browser"
                                            Else
                                                .Hyperlinks.Add Cells(dataRivi, resultStartColumn + 2), "http://www.youtube.com/user/" & accountName    ', "Open channel in browser"
                                            End If
                                        End If
                                        .Cells(dataRivi, resultStartColumn + 1).Font.Bold = True
                                        .Cells(dataRivi, resultStartColumn + 2).Font.Bold = True
                                    End If
                                End If

                                If segmentCount > 1 Then .Cells(dataRivi, resultStartColumn + 3).value = Left$(segmentName, 255)
                            End If


                            stParam1 = "7.61"


                            metricNumResultArr = 1

                            If Left$(arr(1, 1), 6) <> "Error:" And IsArray(arr) Then

                                If UBound(arr, 1) > 1 Then
                                    Debug.Print "WARNING: query type A has result array with more than one row (" & UBound(arr, 1) & ")"
                                    If debugMode = True Then
                                        MsgBox "WARNING: query type A has result array with more than one row (" & UBound(arr, 1) & ")"
                                    End If
                                    For i = 2 To UBound(arr, 1)
                                        For j = 1 To UBound(arr, 2)
                                            If IsNumeric(arr(i, j)) Then
                                                arr(1, j) = arr(1, j) + arr(i, j)
                                            End If
                                        Next j
                                    Next i
                                End If

                                For metricNum = metricSetsArr(metricSetNum, 1) To metricSetsArr(metricSetNum, 2)


                                    '     dataSar = resultStartColumn + 2 + metricNum + doComparisons * (metricNum - 1)

                                    dataSar = findColumnNumber("1|" & metricNum & "|1|1|" & iterationNum)


                                    If iterationNum = 2 Then
                                        columnInfoArr(dataSar, 6) = True
                                    Else
                                        columnInfoArr(dataSar, 6) = False
                                    End If
                                    columnInfoArr(dataSar, 1) = metricsArr(metricNum, 1)
                                    columnInfoArr(dataSar, 2) = metricsArr(metricNum, 8)
                                    columnInfoArr(dataSar, 3) = profName



                                    stParam1 = "7.62"

                                    arvo = vbNullString

                                    If metricsArr(metricNum, 5) <> vbNullString Then

                                        Select Case metricsArr(metricNum, 5)
                                        Case "div"
                                            div = arr(1, metricNumResultArr + 1)
                                            If div > 0 Then
                                                arvo = arr(1, metricNumResultArr) / div
                                            Else
                                                arvo = vbNullString
                                            End If
                                        Case "div*86400"    'avg session
                                            div = arr(1, metricNumResultArr + 1)
                                            If div <> 0 Then
                                                arvo = arr(1, metricNumResultArr) / (86400 * div)
                                            Else
                                                arvo = vbNullString
                                            End If
                                        Case "1000*div"    'CPM
                                            div = arr(1, metricNumResultArr + 1)
                                            If div <> 0 Then
                                                arvo = 1000 * arr(1, metricNumResultArr) / div
                                            Else
                                                arvo = vbNullString
                                            End If
                                        Case "d86400"    'time on site
                                            arvo = arr(1, metricNumResultArr) / 86400
                                        Case "d1000000"    'AW cost, budget
                                            arvo = arr(1, metricNumResultArr) / 1000000
                                        Case "div1000"    'AW CPM
                                            div = arr(1, metricNumResultArr + 1)
                                            If div <> 0 Then
                                                arvo = arr(1, metricNumResultArr) / (1000 * div)
                                            Else
                                                arvo = vbNullString
                                            End If
                                        Case "div1000000"    'AW CPC, cost per conversion
                                            div = arr(1, metricNumResultArr + 1)
                                            If div <> 0 Then
                                                arvo = arr(1, metricNumResultArr) / (1000000 * div)
                                            Else
                                                arvo = vbNullString
                                            End If
                                        Case "div*86400&minus"  'avg time on page
                                            num = arr(1, dimensionsCount + metricNumResultArr + segmDimCount)
                                            div = 86400 * (arr(1, dimensionsCount + metricNumResultArr + 1 + segmDimCount) - arr(1, dimensionsCount + metricNumResultArr + 2 + segmDimCount))
                                            If div <> 0 Then
                                                arvo = num / div
                                            Else
                                                arvo = vbNullString
                                            End If
                                        Case "minus"    'lost impressions
                                            num = arr(arrRivi, dimensionsCountInclSubGlobal + metricNumResultArr + segmDimCount)
                                            div = arr(arrRivi, dimensionsCountInclSubGlobal + metricNumResultArr + 1 + segmDimCount)
                                            If num = vbNullString Then num = 0
                                            If div = vbNullString Then div = 0
                                            arvo = num - div
                                            div = 1
                                            num = arvo
                                        Case "minus&div"
                                            num = arr(arrRivi, dimensionsCountInclSubGlobal + metricNumResultArr + segmDimCount) - arr(arrRivi, dimensionsCountInclSubGlobal + metricNumResultArr + 1 + segmDimCount)
                                            div = arr(arrRivi, dimensionsCountInclSubGlobal + metricNumResultArr + 2 + segmDimCount)
                                            If div <> 0 Then
                                                arvo = num / div
                                            Else
                                                arvo = vbNullString
                                            End If
                                        Case "div&minus&minusone"
                                            num = arr(arrRivi, dimensionsCountInclSubGlobal + metricNumResultArr + segmDimCount)
                                            div = arr(arrRivi, dimensionsCountInclSubGlobal + metricNumResultArr + 1 + segmDimCount) - arr(arrRivi, dimensionsCountInclSubGlobal + metricNumResultArr + 2 + segmDimCount)
                                            If div <> 0 Then
                                                arvo = (num / div) - 1
                                            Else
                                                arvo = vbNullString
                                            End If
                                        Case "div&minus&plus&minusone"
                                            num = arr(arrRivi, dimensionsCountInclSubGlobal + metricNumResultArr + segmDimCount)
                                            div = arr(arrRivi, dimensionsCountInclSubGlobal + metricNumResultArr + 1 + segmDimCount) - arr(arrRivi, dimensionsCountInclSubGlobal + metricNumResultArr + 2 + segmDimCount) + arr(arrRivi, dimensionsCountInclSubGlobal + metricNumResultArr + 3 + segmDimCount)
                                            If div > 0 Then
                                                arvo = (num / div) - 1
                                            Else
                                                arvo = vbNullString
                                            End If
                                        Case Else
                                            arvo = arr(1, metricNumResultArr)
                                        End Select

                                        metricNumResultArr = metricNumResultArr + metricsArr(metricNum, 4)

                                    Else

                                        arvo = arr(1, metricNumResultArr)
                                        metricNumResultArr = metricNumResultArr + 1

                                    End If
                                    stParam1 = "7.63"
                                    If iterationNum = 1 Then
                                        .Cells(dataRivi, dataSar).value = arvo
                                    Else
                                        arvo1 = .Cells(dataRivi, dataSar - 1).value
                                        muutos = vbNullString

                                        Select Case comparisonValueType
                                        Case "perc"
                                            If arvo <> vbNullString And arvo <> 0 Then
                                                If arvo1 <> 0 And arvo1 <> vbNullString Then
                                                    muutos = arvo1 / arvo - 1
                                                End If
                                            ElseIf arvo = 0 Or arvo = vbNullString Then    'both zeroes/blanks
                                                muutos = vbNullString
                                            Else    'newer value zero
                                                muutos = -1
                                            End If
                                        Case "abs"
                                            If arvo = vbNullString Then arvo = 0
                                            If arvo1 = vbNullString Then arvo1 = 0
                                            muutos = arvo1 - arvo
                                        Case "val"
                                            muutos = arvo
                                        End Select

                                        .Cells(dataRivi, dataSar).value = muutos

                                    End If

                                Next metricNum

                                stParam1 = "7.64"
                            Else
                                stParam1 = "7.65"

                                dataSar = vsarData + 1
                                If iterationNum = 1 Then
                                    With .Cells(dataRivi, dataSar)
                                        If arr(1, 1) = "Error: No data found" Then
                                            .value = "No data found"
                                        Else
                                            .value = arr(1, 1)
                                        End If
                                        .Font.Italic = True
                                    End With
                                End If
                            End If

                            'If runningSheetRefresh = False Then
                            Application.ScreenUpdating = True
                            Application.ScreenUpdating = False
                            'End If






                            ' If queriesCompletedCount >= queryCount And foundNonFinishedQuery = False Then Exit Do
                            'If allQueriesFetched = False Then Exit For

                        End If
                    End If
                Else
                    foundNonFinishedQuery = True
                    'Exit For

                End If

                DoEvents

            Next queryNum

            If areAllQueriesPlacedOnSheet() = True Then Exit Do

        Loop
        inDataFetchLoop = False

        stParam1 = "7.70"

        Application.Cursor = xlNormal


        If Not useQTforDataFetch Then Set objHTTPstatus = Nothing




        stParam1 = "7.701"


        Call eraseObjHTTPs

        Call updateProgressIterationBoxes("EXITLOOP")

    End With


    Exit Sub

generalErrHandler:

    stParam2 = "REPORTERROR " & Err.Number & "|" & Err.Description & "|" & Application.StatusBar

    Debug.Print "REPORT ERROR: " & stParam1 & " " & stParam2
    'Call checkE(email, dataSource, True)

    If Err.Number = 18 Then
        Call hideProgressBox
        Call removeTempsheet
        End
    End If

    Resume Next


End Sub



