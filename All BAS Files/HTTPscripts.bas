Attribute VB_Name = "HTTPscripts"
Option Private Module
Option Explicit

Dim objHTTP1 As Object
Dim objHTTP2 As Object
Dim objHTTP3 As Object
Dim objHTTP4 As Object
Dim objHTTP5 As Object
Dim objHTTP6 As Object
Dim objHTTP7 As Object
Dim objHTTP8 As Object
Dim objHTTP9 As Object
Dim objHTTP10 As Object
Dim objHTTP11 As Object
Dim objHTTP12 As Object
Dim objHTTP13 As Object
Dim objHTTP14 As Object
Dim objHTTP15 As Object
Dim objHTTP16 As Object
Dim objHTTP17 As Object
Dim objHTTP18 As Object
Dim objHTTP19 As Object

Sub setMSXML(ByRef obj As Object)
    On Error Resume Next
    Dim i As Integer
    If usingMacOSX = False Then
        For i = 6 To 1 Step -1
            If MSXMLversion = 0 Or MSXMLversion = i Then
                If i = 1 Then
                    Set obj = CreateObject("MSXML2.ServerXMLHTTP")
                Else
                    Set obj = CreateObject("MSXML2.ServerXMLHTTP." & i & ".0")
                End If
                If Not obj Is Nothing Then
                    MSXMLversion = i
                    obj.setOption 2, 13056
                    Exit Sub
                End If
            End If
        Next i
    End If
End Sub

Sub checkArrForErrors()

    On Error Resume Next

    Dim errorStr As String
    Dim errorStrOrig As String
    Dim SDlabelsQueryNum As Long
    Dim queryNumLoc As Long
    Dim arr2
    Dim reloginNowResp As Long

    If Not IsArray(arr) Then
        Debug.Print "Invalid array, query " & queryNum
        errorStr = arr
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = errorStr
    End If

    If UBound(arr, 1) = 1 And (Left$(arr(1, 1), 6) = "Error:" Or InStr(1, arr(1, 1), "500 Internal Server Error") > 0) Then
        errorStrOrig = arr(1, 1)
        errorStr = LCase(arr(1, 1))
        Debug.Print "Query error: " & errorStr
    Else
        Exit Sub
    End If


    If InStr(1, errorStr, "no data found") > 0 Then
        '  If (usingMacOSX = True And queryArr(queryNum, 16) <= 1) Or (usingMacOSX = False And queryArr(queryNum, 16) <= 0) Then
        If (usingMacOSX = False And queryArr(queryNum, 16) < 1) Then
            SDlabelsQuery = queryArr(queryNum, 3)
            SDlabelsQueryNum = queryArr(queryNum, 12)
            Call setQueryAsNotFetched
            queryArr(queryNum, 16) = queryArr(queryNum, 16) + 1   'error count
            queryArr(queryNum, 12) = SDlabelsQueryNum
            Debug.Print "No data in query " & queryNum & ", rerunning.. (error count: " & queryArr(queryNum, 16) & ")"
        End If
        Exit Sub
    ElseIf InStr(1, errorStr, "500 internal server error") > 0 Then
        If (avoidSampling Or usingMacOSX) And queryArr(queryNum, 16) <= 2 Then     'Or (usingMacOSX = False And queryArr(queryNum, 16) <= 0) Then
            SDlabelsQuery = queryArr(queryNum, 3)
            SDlabelsQueryNum = queryArr(queryNum, 12)
            avoidSampling = False
            Call setQueryAsNotFetched
            queryArr(queryNum, 16) = queryArr(queryNum, 16) + 1   'error count
            queryArr(queryNum, 12) = SDlabelsQueryNum
            Debug.Print "500 Internal Server Error in query " & queryNum & ", rerunning.. (error count: " & queryArr(queryNum, 16) & ")"
        Else
            arr(1, 1) = "Error: server error. These are usually caused by queries returning too much data."
            queryArr(queryNum, 8) = arr
        End If
    ElseIf (dataSource = "FB" Or dataSource = "FA") And InStr(1, errorStr, "access token expired") > 0 Then
        reloginNowResp = MsgBox("Facebook access token has expired, you need to renew Supermetrics Data Grabber's access rights to your data. Do you wish to do so now?", vbYesNo, "Renew Facebook access rights")
        If reloginNowResp = vbYes Then
            email = getEmailForToken(authToken, True)

            If updatingPreviouslyCreatedSheet = False Then
                'saves current inteface query
                Call markToCurrentQuery
                Call saveQueryFromCQ
            End If

            Call logoutOneAccount(email)
            If Range("loggedin" & varsuffix).value = True Then
                If dataSource = "FB" Then
                    Call addLoginFB
                Else
                    Call addLoginFA
                End If
            Else
                If dataSource = "FB" Then
                    Call OAuthLoginFB
                Else
                    Call OAuthLoginFA
                End If
            End If
            If updatingPreviouslyCreatedSheet = True Then
                dataSheet.Select
                Call refreshDataOnSelectedSheet
            Else
                Call returnSavedQueryToCQ
                Call getFromCurrentQuery
                Call runReportFB
            End If
            End
        Else
            Debug.Print "FB token expired"
        End If
    ElseIf InStr(1, errorStr, "license expired") > 0 Or InStr(1, errorStr, "demo expired") > 0 Then
        email = queryArr(queryNum, 18)
        With buyLicenseBox
            .note1.Caption = "Your Supermetrics Data Grabber " & moduleName & " license for account " & email & " has expired. Please visit Supermetrics.com to purchase a license."
            Call hideProgressBox
            Call removeTempsheet
            If updatingPreviouslyCreatedSheet = False Then Call removeDatasheet
            .Show
            End
        End With
    ElseIf InStr(1, errorStr, "insufficient quota to proceed") > 0 Or InStr(1, errorStr, "rate limit exceeded") > 0 Then
        Debug.Print "Quota exceeded, rerunning querynum " & queryNum & " profnum " & profNum
        SDlabelsQuery = queryArr(queryNum, 3)
        SDlabelsQueryNum = queryArr(queryNum, 12)
        Call setQueryAsNotFetched
        queryArr(queryNum, 16) = queryArr(queryNum, 16) + 1   'error count
        queryArr(queryNum, 12) = SDlabelsQueryNum
    ElseIf (InStr(1, errorStr, "token") > 0 Or InStr(1, errorStr, "authenticat") > 0 Or InStr(1, errorStr, "unauthorized") > 0 Or InStr(1, errorStr, "401") > 0) And tokenRefreshed = False Then
        If dataSource = "GA" And clientLoginModeForGA = True Then
            MsgBox "Your authentication has expired. Please retransfer the profile list to Supermetrics Data Grabber."
        Else
            Debug.Print "Authentication error, querynum " & queryNum & ", rerunning all queries"
            authToken = refreshToken(authToken)
            If authToken <> vbNullString Then
                Range("authToken" & varsuffix).value = authToken
                Call storeTokenToSheet("ALL", authToken, email)
            End If
            '   End If
            tokenRefreshed = True
            Call initializeFetchArrays
        End If
        Exit Sub
    ElseIf InStr(1, errorStr, "token") > 0 Or InStr(1, errorStr, "authenticat") > 0 Or InStr(1, errorStr, "unauthorized") > 0 Then
        Call hideProgressBox
        Application.StatusBar = False
        MsgBox "Authentication with " & serviceName & " failed. Please check your login information. This may also be due to problems with the network connection, or temporary problems with " & serviceName & "." & vbCrLf & vbCrLf & "The error message is: " & errorStrOrig
        End
    ElseIf InStr(1, errorStr, "IllegalDimensionMetricCombinationException") > 0 Then
        On Error Resume Next
        Application.DisplayAlerts = False
        If runningSheetRefresh = False Then Call removeSheet
        Application.DisplayAlerts = False
        tempSheet.Delete
        Call hideProgressBox
        Analytics.Select
        Application.DisplayAlerts = True
        Application.StatusBar = False
        If segmentID = vbNullString And filterStr = vbNullString Then
            MsgBox "The selected combination of metrics and dimensions is not allowed by the " & serviceName & " API." & vbCrLf & vbCrLf & errorStr, , "Illegal field combination"
        Else
            MsgBox "The selected combination of metrics and dimensions is not allowed by the " & serviceName & " API. Note that filter or segments can also cause illegal combinations." & vbCrLf & vbCrLf & errorStrOrig, , "Illegal field combination"
        End If
        Exit Sub
    ElseIf InStr(1, errorStr, "!!!") > 0 Then
        queryArr(queryNum, 16) = queryArr(queryNum, 16) + 1   'error count
        If queryArr(queryNum, 16) < 2 Then
            Debug.Print "Rerunning querynum " & queryNum & " profnum " & profNum
            SDlabelsQuery = queryArr(queryNum, 3)
            SDlabelsQueryNum = queryArr(queryNum, 12)
            Call setQueryAsNotFetched
            queryArr(queryNum, 16) = queryArr(queryNum, 16) + 1   'error count
            queryArr(queryNum, 14) = queryArr(queryNum, 14) & "%createNewDefinition->TRUE%"
            queryArr(queryNum, 12) = SDlabelsQueryNum
        Else
            Debug.Print "Error count exceeded on query " & queryNum & ", not rerunning..  " & errorStr
        End If
    End If




End Sub

Public Function areAllQueriesPlacedOnSheet()
    Dim queryNumLoc As Long
    Dim foundNonCompleteQuery As Boolean
    foundNonCompleteQuery = False
    For queryNumLoc = 1 To queryCount
        If queryArr(queryNumLoc, 10) <> True Then
            foundNonCompleteQuery = True
            Exit For
        End If
    Next queryNumLoc
    areAllQueriesPlacedOnSheet = Not foundNonCompleteQuery
End Function





Sub runQueriesOnFreeObjHTTPs()

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim errorStr As String
    Dim maxObjHTTPnum As Integer

    Dim freeObjHTTPs As Boolean

    Dim checkNum As Integer
    Dim parameterValue As Long

    Dim prevProfID As String
    Dim settingsParameterAll As String
    Dim settingsParameter As String

    Dim dataSourceOrig As String
    Dim processID As Long
    Dim profStr As String

    Dim maxResultsLoc As Long

    settingsParameterAll = vbNullString
    If sumAllProfiles Then settingsParameterAll = rscL2 & "SUM_ALL"
    If rawDataReport And segmentCount > 1 Then settingsParameterAll = settingsParameterAll & rscL2 & "INCLUDE_SEGMENT"
    If advancedSettings <> "" Then settingsParameterAll = settingsParameterAll & rscL2 & advancedSettings
    settingsParameter = settingsParameterAll

    allQueriesStarted = True

    maxObjHTTPnum = maxSimultaneousQueries
    If maxSimultaneousQueries > queryCount Then maxObjHTTPnum = queryCount

    For checkNum = 1 To 10
        prevProfID = vbNullString
        For queryNum = 1 To queryCount
            If queryArr(queryNum, 6) = False Then    'query not yet running
                allQueriesStarted = False
                profNum = queryArr(queryNum, 1)
                profID = queryArr(queryNum, 2)

                If (profID <> prevProfID And profilesArr(profNum, 6) <= checkNum) Or usingMacOSX Or useQTforDataFetch Then

                    'check if free objHTTP available
                    freeObjHTTPs = False
                    For objHTTPnum = 1 To maxSimultaneousQueries
                        If objHTTParr(objHTTPnum, 1) = False Or (objHTTPnum = 1 And (usingMacOSX Or useQTforDataFetch)) Then
                            If objHTTPnum < maxSimultaneousQueries Then freeObjHTTPs = True
                            Debug.Print "RUNNING fetching data: querynum=" & queryNum & " objhttpnum=" & objHTTPnum
                            '  Call updateProgressAdditionalMessage("Started fetching data for query " & queryNum & "/" & queryCount)
                            If usingMacOSX = False And useQTforDataFetch = False Then
                                objHTTParr(objHTTPnum, 1) = True
                                objHTTParr(objHTTPnum, 2) = queryNum
                                queryArr(queryNum, 6) = True  'for QT queries, mark as running only when queryIDforDB obtained
                            End If

                            queryArr(queryNum, 7) = objHTTPnum
                            profNum = queryArr(queryNum, 1)
                            profID = queryArr(queryNum, 2)
                            SDlabelsQuery = queryArr(queryNum, 3)
                            SDothersQuery = queryArr(queryNum, 22)
                            profilesArr(profNum, 6) = profilesArr(profNum, 6) + 1  'queries running
                            profStr = profilesArr(profNum, 8)
                            prevProfID = profID

                            dataSourceOrig = dataSource

                            '   dataSource = metricSetsArr(queryArr(queryNum, 19), 8)
                            '   If dataSourceOrig <> dataSource Then Call setDatasourceVariables

                            If dataSourceOrig = "GA" Or dataSourceOrig <> dataSource Then
                                metrics = metricSetsArr(queryArr(queryNum, 19), 4)
                                dimensions = metricSetsArr(queryArr(queryNum, 19), 6)
                                sortStr = metricSetsArr(queryArr(queryNum, 19), 7)
                            End If

                            If segmentIsAllVisits = False Or (rawDataReport And segmentCount > 1) Then
                                segmentNum = queryArr(queryNum, 21)
                                segmentID = segmentArr(segmentNum, 1)
                                segmentName = segmentArr(segmentNum, 2)
                            Else
                                segmentNum = 1
                                segmentID = ""
                                segmentName = "All Visits"
                            End If


                            If queryArr(queryNum, 4) = 1 Then  'iterationNum = 1
                                startDate = startDate1
                                endDate = endDate1
                                If dataSource = "GA" Then maxResultsLoc = maxResults1
                            Else
                                startDate = startDate2
                                endDate = endDate2
                                If dataSource = "GA" Then maxResultsLoc = maxResults1 * maxResultsMultiplierForComparisonQuery
                            End If

                            If dataSource = "GA" Then
                                HTTPrequestType = "POST"
                                HTTPpostStr = ""
                                authToken = getTokenFromSheet(profID)
                                If authToken = vbNullString Then
                                    Debug.Print "Empty authtoken"
                                End If

                                If queryArr(queryNum, 23) And maxResultsLoc < 500000 Then maxResultsLoc = 500000

                                If queryType = "A" Then
                                    HTTPpostStr = createGetGAdataURL(profStr, metrics, startDate, endDate, filterStr, dimensions, segmentID, , maxResults)
                                ElseIf SDlabelsQuery Then    'sdlabelsquery
                                    If allProfilesInOneQuery Then
                                        maxResultsLoc = 2000
                                    ElseIf dimensionsRequiringCompressionInSD Then
                                        maxResultsLoc = 500000
                                    Else
                                        maxResultsLoc = segmDimCategoriesCount - 1
                                    End If
                                    If segmDimIsTime = False Or segmDimHasNonTimeComponent = True Then
                                        HTTPpostStr = createGetGAdataURL(profStr, firstMetricStr, startDate, endDate, filterStr, segmDimNameComb, segmentID, sortStr, maxResultsLoc)
                                    Else
                                        HTTPpostStr = createGetGAdataURL(profStr, firstMetricStr, startDate, endDate, filterStr, segmDimNameComb, segmentID, , maxResultsLoc)
                                    End If
                                ElseIf SDothersQuery Then
                                    HTTPpostStr = createGetGAdataURL(profStr, metrics, startDate, endDate, filterStr, dimensionsBasicStr, segmentID, sortStr, maxResultsLoc)
                                Else
                                    HTTPpostStr = createGetGAdataURL(profStr, metrics, startDate, endDate, filterStr, dimensions, segmentID, sortStr, maxResultsLoc)
                                End If


                                HTTPpostStr = HTTPpostStr & "&token=" & uriEncode(authToken)
                                HTTPpostStr = HTTPpostStr & "&includeHeaders=0"
                                HTTPpostStr = HTTPpostStr & "&includeLS=0"

                                If Not SDlabelsQuery And avoidSampling Then settingsParameter = settingsParameterAll & rscL2 & "AVOID_SAMPLING"

                            Else
                                ' dataSource = "AW" "YT" "AC" "FL" "FB"

                                HTTPrequestType = "POST"
                                authToken = getTokenFromSheet(profID)

                                If allProfilesInOneQuery Then
                                    HTTPpostStr = "profiles=" & uriEncode(allProfilesStr)
                                Else
                                    HTTPpostStr = "profiles=" & uriEncode(profStr)
                                End If

                                HTTPpostStr = HTTPpostStr & "&start-date=" & uriEncode(Year(startDate) & "-" & Format(Month(startDate), "00") & "-" & Format(Day(startDate), "00"))
                                HTTPpostStr = HTTPpostStr & "&end-date=" & uriEncode(Year(endDate) & "-" & Format(Month(endDate), "00") & "-" & Format(Day(endDate), "00"))

                                If filterStr <> "" Then HTTPpostStr = HTTPpostStr & "&filterStr=" & uriEncode(filterStr)

                                If queryType <> "A" Then
                                    If SDlabelsQuery Then
                                        HTTPpostStr = HTTPpostStr & "&dimensions=" & uriEncode(segmDimNameComb)
                                        HTTPpostStr = HTTPpostStr & "&metrics=" & uriEncode(firstMetricStr)
                                    ElseIf SDothersQuery Then
                                        HTTPpostStr = HTTPpostStr & "&dimensions=" & uriEncode(dimensionsBasicStr)
                                        HTTPpostStr = HTTPpostStr & "&metrics=" & uriEncode(metrics)
                                    Else
                                        HTTPpostStr = HTTPpostStr & "&dimensions=" & uriEncode(dimensions)
                                        HTTPpostStr = HTTPpostStr & "&metrics=" & uriEncode(metrics)
                                    End If
                                Else
                                    HTTPpostStr = HTTPpostStr & "&metrics=" & uriEncode(metrics)
                                End If

                                HTTPpostStr = HTTPpostStr & "&useOAuth=1"
                                If usingMacOSX Then HTTPpostStr = HTTPpostStr & "&chrEncode=1&chrEncodeType=2"
                                HTTPpostStr = HTTPpostStr & "&token=" & uriEncode(authToken)
                                HTTPpostStr = HTTPpostStr & "&includeHeaders=0"
                                HTTPpostStr = HTTPpostStr & "&includeLS=0"


                            End If

                            If nameEncodingStr <> vbNullString Then HTTPpostStr = HTTPpostStr & "&encodings=" & uriEncode(nameEncodingStr)

                            HTTPpostStr = HTTPpostStr & "&email=" & uriEncode(email)
                            HTTPpostStr = HTTPpostStr & "&appid=" & appID
                            HTTPpostStr = HTTPpostStr & "&dataSource=" & dataSource
                            HTTPpostStr = HTTPpostStr & "&version=" & uriEncode(versionNumber)
                            HTTPpostStr = HTTPpostStr & "&system=" & uriEncode(OSandExcelVersion)
                            HTTPpostStr = HTTPpostStr & "&rid=" & randID

                            If SDlabelsQuery Then HTTPpostStr = HTTPpostStr & "&sd=1"




                            If separatorList = vbNullString Then
                                separatorList = "&rscL1=" & uriEncode(rscL1)
                                separatorList = separatorList & "&rscL2=" & uriEncode(rscL2)
                                separatorList = separatorList & "&rscL3=" & uriEncode(rscL3)
                                separatorList = separatorList & "&rscL4=" & uriEncode(rscL4)
                                If useQTforDataFetch Then separatorList = separatorList & "&rscL0=" & uriEncode(rscL0)
                            End If

                            HTTPpostStr = HTTPpostStr & separatorList

                            If usingMacOSX = True Or useQTforDataFetch = True Then
                                HTTPpostStr = HTTPpostStr & "&separateFetchQuery=TRUE"
                                HTTPpostStr = HTTPpostStr & "&chrencode=true"
                                HTTPpostStr = HTTPpostStr & "&urlencode=true"
                            End If




                            processID = Round(1000000000 * Rnd, 0)
                            HTTPpostStr = HTTPpostStr & "&pid=" & processID

                            processIDsStr = processIDsStr & "," & processID


                            If settingsParameter <> vbNullString Then HTTPpostStr = HTTPpostStr & "&settings=" & uriEncode(settingsParameter)


                            queryURL = "https://supermetrics.com/api/getData?responseFormat=RSCL"


                            'starts the query
                            Call runObjHTTP


                            dataSource = dataSourceOrig
                            Call setDatasourceVariables

                            Exit For
                        End If
                    Next objHTTPnum



                    If freeObjHTTPs = False Then Exit Sub
                End If
            End If
        Next queryNum
        If allQueriesStarted Then Exit Sub
        If usingMacOSX Or useQTforDataFetch Then Exit For
    Next checkNum

    If combinedRequestStr <> vbNullString And (usingMacOSX Or useQTforDataFetch) Then
        Dim tempArr As Variant
        Dim i As Long
        Dim queryNumLoc2 As Long
        Dim queryIDforDB As String
        Dim combinedRequestStrEncodings As String
        Dim encodingNum As Integer
        Dim strToEncode As String

        combinedRequestStr = combinedRequestStr & "&appid=" & appID
        combinedRequestStr = combinedRequestStr & "&version=" & uriEncode(versionNumber)
        combinedRequestStr = combinedRequestStr & "&system=" & uriEncode(OSandExcelVersion)
        combinedRequestStr = combinedRequestStr & "&rid=" & randID
        combinedRequestStr = combinedRequestStr & "&rscL2=" & uriEncode(rscL2)
        combinedRequestStrEncodings = ""


        For encodingNum = 1 To 6
            Select Case encodingNum
            Case 1
                strToEncode = "&appid=" & appID & "&dataSource=" & dataSource & "&version=" & uriEncode(versionNumber) & "&system=" & uriEncode(OSandExcelVersion) & "&rid=" & randID
            Case 2
                strToEncode = separatorList & "&separateFetchQuery=TRUE" & "&chrencode=true" & "&urlencode=true"
            Case 3
                strToEncode = "&token=" & uriEncode(authToken)
            Case 4
                strToEncode = "&includeHeaders=0" & "&includeLS=0"
            Case 5
                strToEncode = "&start-date=" & uriEncode(Year(startDate) & "-" & Format(Month(startDate), "00") & "-" & Format(Day(startDate), "00")) & "&end-date=" & uriEncode(Year(endDate) & "-" & Format(Month(endDate), "00") & "-" & Format(Day(endDate), "00"))
            Case 6
                strToEncode = "&dimensions=" & uriEncode(dimensions) & "&metrics=" & uriEncode(metrics)
            End Select
            If Len(Replace(combinedRequestStr, uriEncode(strToEncode), rscL3 & encodingNum & rscL3)) < Len(combinedRequestStr) Then
                combinedRequestStrEncodings = combinedRequestStrEncodings & rscL1 & encodingNum & rscL2 & strToEncode
                combinedRequestStr = Replace(combinedRequestStr, uriEncode(strToEncode), rscL3 & encodingNum & rscL3)
            End If
        Next encodingNum

        If combinedRequestStrEncodings <> vbNullString Then combinedRequestStr = combinedRequestStr & "&encodings=" & uriEncode(combinedRequestStrEncodings) & "&rscL1=" & uriEncode(rscL1) & "&rscL3=" & uriEncode(rscL3)

        If debugMode Then Debug.Print "combinedRequestStr: " & combinedRequestStr
        If debugMode Then Debug.Print "LEN combinedRequestStr: " & Len(combinedRequestStr)

        queryURL = "https://supermetrics.com/api/getDataMulti?responseFormat=RSCL"
        queryTableResultStr = ""
        Call fetchDataWithQueryTableDirect(queryURL, combinedRequestStr)
        tempArr = Split(queryTableResultStr, "#")
        For i = 0 To UBound(tempArr)
            queryNumLoc2 = parseVarFromStr(tempArr(i), "QN", rscL2)
            queryIDforDB = parseVarFromStr(tempArr(i), "queryIDforDB", rscL2)
            If queryIDforDB = vbNullString Then queryIDforDB = Split(tempArr(i), "|")(1)
            If queryIDforDB = vbNullString Or InStr(1, queryIDforDB, "ERROR") > 0 Then

            Else
                queryArr(queryNumLoc2, 6) = True
                queryArr(queryNumLoc2, 20) = queryIDforDB
                If debugMode = True Then Debug.Print "queryIDforDB for " & queryNumLoc2 & ": " & queryIDforDB
            End If
        Next i
        queryTableResultStr = ""
        combinedRequestStr = ""
    End If

End Sub



Sub runObjHTTP()

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim SDlabelsQueryNum As Long
    Dim queryIDforDB As String




    If usingMacOSX = False And useQTforDataFetch = False Then
        Call assignObjHTTP

        If HTTPrequestType = "POST" Then
            objhttp.Open "POST", queryURL, True
            If debugMode = True Then Debug.Print "HTTPpostStr Q" & queryNum & ": " & HTTPpostStr
        Else
            objhttp.Open "GET", queryURL, True
            If debugMode = True Then Debug.Print "queryURL Q" & queryNum & ": " & queryURL
        End If

        objhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"

        objhttp.setTimeouts 100000000, 100000000, 100000000, 100000000
        objhttp.setOption 2, 13056

        On Error GoTo errhandler
        If HTTPrequestType = "POST" Then
            objhttp.send (HTTPpostStr)
        Else
            objhttp.send (vbNullString)
        End If
        On Error Resume Next
        If debugMode = True Then On Error GoTo 0

    Else
        If Len(combinedRequestStr) < 20000 Then

            If debugMode = True Then Debug.Print "HTTPpostStr Q" & queryNum & ": " & HTTPpostStr

            If combinedRequestStr = vbNullString Then
                combinedRequestStr = "query" & queryNum & "=" & uriEncode(HTTPpostStr)
            Else
                combinedRequestStr = combinedRequestStr & "&query" & queryNum & "=" & uriEncode(HTTPpostStr)
            End If
        Else

            If debugMode Then Debug.Print "to next one"

        End If

    End If

    Exit Sub

errhandler:
    Debug.Print "Error: sending query " & queryNum & " failed, rerunning..."
    SDlabelsQueryNum = queryArr(queryNum, 12)
    SDlabelsQuery = queryArr(queryNum, 3)
    Call setQueryAsNotFetched
    queryArr(queryNum, 12) = SDlabelsQueryNum
    objHTTPnum = queryArr(queryNum, 7)
    objHTTParr(objHTTPnum, 1) = False
    objHTTParr(objHTTPnum, 2) = ""
    Exit Sub

End Sub



Sub checkForCompletedObjHTTPs()


    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim warningText As String
    Dim SDlabelsQueryNum As Long
    Dim arr As Variant
    Dim requestStr As String
    Dim responseStr As String

    If allQueriesStarted = True Then allQueriesFetched = True

    If usingMacOSX = False And useQTforDataFetch = False Then

        For objHTTPnum = 1 To maxSimultaneousQueries
            If objHTTParr(objHTTPnum, 1) = True And objHTTParr(objHTTPnum, 2) <> "" Then
                queryReady = False
                Call assignObjHTTP
                ' Debug.Print "Checking if query finished, objhhtpnum=" & objHTTPnum
                If objhttp.readyState = 4 Then
                    queryReady = True
                    queryNum = objHTTParr(objHTTPnum, 2)
                    On Error GoTo reRunQuery
                    If dataSource = "AC" And usingMacOSX = False Then
                        queryArr(queryNum, 8) = chrDecode(objhttp.responsetext)
                    Else
                        queryArr(queryNum, 8) = objhttp.responsetext
                    End If
                    On Error Resume Next
                    queryArr(queryNum, 7) = ""  ' objHTTPnum where running
                    If debugMode = True Then Debug.Print "Response: " & Left(queryArr(queryNum, 8), 2000)
                Else
                    queryNum = objHTTParr(objHTTPnum, 2)
                    '  updateProgressAdditionalMessage ("Waiting for query " & queryNum & "/" & queryCount & " to finish")
                    On Error GoTo reRunQuery
                    objhttp.waitForResponse 0
                    On Error Resume Next
                    If debugMode = True Then On Error GoTo 0
                    allQueriesFetched = False
                End If


                If queryReady = True Then

                    objHTTParr(objHTTPnum, 1) = False
                    queryArr(queryNum, 9) = True
                    Debug.Print "Query finished, querynum=" & queryNum & " objhhtpnum=" & objHTTPnum
                    '  Call updateProgressAdditionalMessage("Data fetched for query " & queryNum & "/" & queryCount)

                    profNum = queryArr(queryNum, 1)
                    profilesArr(profNum, 6) = profilesArr(profNum, 6) - 1

                    'If debugMode = True Then Debug.Print "GARESPONSE: " & queryArr(queryNum, 8)

                    If InStr(1, queryArr(queryNum, 8), "IllegalDimensionMetricCombinationException") > 0 Then
                        Application.DisplayAlerts = False

                        If runningSheetRefresh = False Then Call removeSheet
                        Application.DisplayAlerts = False
                        tempSheet.Delete
                        Call hideProgressBox
                        configsheet.Select
                        Application.DisplayAlerts = True
                        Application.StatusBar = False

                        warningText = "The selected combination of metrics and dimensions is not allowed by the " & serviceName & " API."
                        If segmentID <> vbNullString Or filterStr <> vbNullString Then warningText = warningText & " Note that the filter string or segments can also cause illegal combinations."
                        warningText = warningText & vbCrLf & vbCrLf & parseVarFromStr(queryArr(queryNum, 8), "ERROR", rscL2)

                        MsgBox warningText
                        End
                    End If

                End If

            End If
        Next objHTTPnum

    Else
        For queryNum = 1 To queryCount
            If queryArr(queryNum, 9) = False And queryArr(queryNum, 20) <> vbNullString Then
                requestStr = "queryIDforDB=" & queryArr(queryNum, 20)
                requestStr = requestStr & "&acceptStatus=TRUE&pid=" & uriEncode(processIDsStr) & "&type=multi"
                requestStr = requestStr & "&appid=" & appID & "&version=" & versionNumber
                requestStr = requestStr & "&rscL0=" & uriEncode(rscL0)
                requestStr = requestStr & "&rscL1=" & uriEncode(rscL1)
                stParam1 = "3.054"
                Call fetchDataWithQueryTableDirect("https://supermetrics.com/api/getQueryResults?responseFormat=RSCL", requestStr, True, True)
                responseStr = queryTableResultStr
                If Left(responseStr, 1) = "%" Then
                    responseStr = URL_decode(responseStr)
                    responseStr = UTF8_Decode(responseStr)
                End If
                If Left(responseStr, 1) = "%" Then
                    responseStr = URL_decode(responseStr)
                    responseStr = UTF8_Decode(responseStr)
                End If
                If parseVarFromStr(responseStr, "RESULTS", rscL0) <> vbNullString Then
                    queryArr(queryNum, 8) = parseVarFromStr(responseStr, "RESULTS", rscL0)
                    queryReady = True
                    Debug.Print "Query finished, querynum=" & queryNum
                    If debugMode = True Then Debug.Print "Response: " & Left(queryArr(queryNum, 8), 2000)
                    queryArr(queryNum, 9) = True
                    '  Call updateProgressAdditionalMessage("Data fetched for query " & queryNum & "/" & queryCount)
                    If InStr(1, queryArr(queryNum, 8), "IllegalDimensionMetricCombinationException") > 0 Then
                        Application.DisplayAlerts = False

                        If runningSheetRefresh = False Then Call removeSheet
                        Application.DisplayAlerts = False
                        If queryType = "SD" Then tempSheet.Delete
                        Call hideProgressBox
                        configsheet.Select
                        Application.DisplayAlerts = True
                        Application.StatusBar = False

                        warningText = "The selected combination of metrics and dimensions is not allowed by the " & serviceName & " API."
                        If segmentID <> vbNullString Or filterStr <> vbNullString Then warningText = warningText & " Note that the filter string or segments can also cause illegal combinations."
                        warningText = warningText & vbCrLf & vbCrLf & parseVarFromStr(queryArr(queryNum, 8), "ERROR", rscL2)

                        MsgBox warningText
                        End
                    End If
                Else
                    DoEvents
                    '  updateProgressAdditionalMessage ("Waiting for query " & queryNum & "/" & queryCount & " to finish")
                    allQueriesFetched = False
                End If
                If parseVarFromStr(responseStr, "QUERIESTOTAL", rscL1) <> vbNullString Then
                    processQueriesCompleted = CLng(parseVarFromStr(responseStr, "QUERIESDONE", rscL1))
                    processQueriesTotal = CLng(parseVarFromStr(responseStr, "QUERIESTOTAL", rscL1))
                End If
                Debug.Print responseStr
            ElseIf queryArr(queryNum, 20) = vbNullString Then
                '  updateProgressAdditionalMessage ("Waiting for query " & queryNum & "/" & queryCount & " to finish")
                allQueriesFetched = False
            End If
        Next queryNum
    End If



    Exit Sub

reRunQuery:
    If queryArr(queryNum, 16) < 5 Then
        Debug.Print "Error on query " & queryNum & ": unrecognized response. Rerunning..."
        queryReady = False
        SDlabelsQuery = queryArr(queryNum, 3)
        SDlabelsQueryNum = queryArr(queryNum, 12)
        querynNumOfPrevSDlabelsQuery = SDlabelsQueryNum
        profNum = queryArr(queryNum, 1)
        iterationNum = queryArr(queryNum, 4)
        Call setQueryAsNotFetched
        queryArr(queryNum, 12) = SDlabelsQueryNum
        queryArr(queryNum, 16) = queryArr(queryNum, 16) + 1
        objHTTPnum = queryArr(queryNum, 7)
        objHTTParr(objHTTPnum, 1) = False
        objHTTParr(objHTTPnum, 2) = ""
    Else
        Debug.Print "Error on query " & queryNum & ": unrecognized response. All reruns done, will return error result..."
        Call testConnection
        queryReady = True
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = "Error: Data could not be fetched " & Err.Description
        queryArr(queryNum, 8) = arr
        queryArr(queryNum, 9) = True
        queryArr(queryNum, 11) = True
        objHTTPnum = queryArr(queryNum, 7)
        objHTTParr(objHTTPnum, 1) = False
        objHTTParr(objHTTPnum, 2) = ""
        Exit Sub
    End If
End Sub

Sub setQueryAsNotFetched(Optional initialQuerySetup As Boolean = False)

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    If queryNum > queryCount Then Exit Sub
    ' If profID = vbNullString Then profID = profilesArr(profNum, 3)
    profID = profilesArr(profNum, 3)

    If SDlabelsQuery = True Then
        If initialQuerySetup = True Then queryArr(queryNum, 1) = profNum
        If initialQuerySetup = True Then queryArr(queryNum, 2) = profID
        If initialQuerySetup = True Then queryArr(queryNum, 3) = True
        If initialQuerySetup = True Then queryArr(queryNum, 4) = 1
        If initialQuerySetup = True Then queryArr(queryNum, 5) = 1
        queryArr(queryNum, 6) = False
        queryArr(queryNum, 9) = False
        queryArr(queryNum, 10) = False
        queryArr(queryNum, 11) = False
        If initialQuerySetup = True Then queryArr(queryNum, 12) = queryNum
        If initialQuerySetup = True Then queryArr(queryNum, 18) = profilesArr(profNum, 4)     'email
        If initialQuerySetup = True Then queryArr(queryNum, 19) = 1
        If initialQuerySetup = True Then queryArr(queryNum, 21) = segmentNum
    Else
        If initialQuerySetup = True Then queryArr(queryNum, 1) = profNum
        If initialQuerySetup = True Then queryArr(queryNum, 2) = profID
        If initialQuerySetup = True Then queryArr(queryNum, 3) = False
        If initialQuerySetup = True Then queryArr(queryNum, 4) = iterationNum
        If initialQuerySetup = True Then queryArr(queryNum, 5) = 1
        queryArr(queryNum, 6) = False
        queryArr(queryNum, 9) = False
        queryArr(queryNum, 10) = False
        queryArr(queryNum, 11) = False
        If initialQuerySetup = True Then queryArr(queryNum, 12) = querynNumOfPrevSDlabelsQuery
        If initialQuerySetup = True Then queryArr(queryNum, 18) = profilesArr(profNum, 4)    'email
        If initialQuerySetup = True Then queryArr(queryNum, 19) = metricSetNum
        If initialQuerySetup = True Then queryArr(queryNum, 21) = segmentNum
        If initialQuerySetup = True Then queryArr(queryNum, 22) = SDothersQuery
        If initialQuerySetup = True And metricSetsArr(metricSetNum, 10) Then queryArr(queryNum, 23) = True
    End If


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

    allQueriesStarted = False
    allQueriesFetched = False
    allXMLsParsed = False


End Sub





Sub initializeFetchArrays()

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Debug.Print "Initializing queryarr"

    profID = vbNullString
    queryNum = 0
    querynNumOfPrevSDlabelsQuery = 0
    SDothersQuery = False
    'populate queryArr
    For profNum = 1 To profileCount
        For segmentNum = 1 To segmentCount
            If queryType = "SD" Then
                queryNum = queryNum + 1
                SDlabelsQuery = True
                Call setQueryAsNotFetched(True)
                queryArr(queryNum, 16) = 0  'error count
                querynNumOfPrevSDlabelsQuery = queryNum
            End If
            SDlabelsQuery = False

            For iterationNum = 1 To iterationsCount
                For metricSetNum = 1 To metricSetsCount
                    queryNum = queryNum + 1
                    '  profID = metricSetsArr(metricSetNum, 9)
                    Call setQueryAsNotFetched(True)
                    queryArr(queryNum, 16) = 0  'error count
                    queryArr(queryNum, 17) = queryNum  'parent query where subquerynum = 1

                    If queryType = "SD" And includeOther Then
                        queryNum = queryNum + 1
                        SDothersQuery = True
                        Call setQueryAsNotFetched(True)
                        queryArr(queryNum, 16) = 0  'error count
                        SDothersQuery = False
                    End If
                    SDlabelsQuery = False

                Next metricSetNum
            Next iterationNum
        Next segmentNum
    Next profNum

    maxParameterValue = iterationsCount
    If metricSetsCount > maxParameterValue Then maxParameterValue = metricSetsCount
    If segmentCount > maxParameterValue Then maxParameterValue = segmentCount


    For objHTTPnum = 1 To maxSimultaneousQueries
        objHTTParr(objHTTPnum, 1) = False
        objHTTParr(objHTTPnum, 2) = ""
    Next objHTTPnum

    allQueriesStarted = False
    allQueriesFetched = False
    allXMLsParsed = False


    queryCount = queryNum
    For queryNum = 1 To queryCount
        Debug.Print "Q " & queryNum & "  prof: " & queryArr(queryNum, 1) & "  SD: " & queryArr(queryNum, 3) & "  ite: " & queryArr(queryNum, 4) & "  subq: " & queryArr(queryNum, 5) & "  segm: " & queryArr(queryNum, 21) & " qwithsdlabels: " & queryArr(queryNum, 12)
    Next queryNum

    'for querynum = 1 to querycount : Print "Q " & queryNum & "  prof: " & queryArr(queryNum, 1) & "  SD: " & queryArr(queryNum, 3) & "  ite: " & queryArr(queryNum, 4) & "  subq: " & queryArr(queryNum, 5) : next querynum

End Sub



Sub assignObjHTTP()
    Set objhttp = Nothing
    Select Case objHTTPnum
    Case 1
        Set objhttp = objHTTP1
    Case 2
        Set objhttp = objHTTP2
    Case 3
        Set objhttp = objHTTP3
    Case 4
        Set objhttp = objHTTP4
    Case 5
        Set objhttp = objHTTP5
    Case 6
        Set objhttp = objHTTP6
    Case 7
        Set objhttp = objHTTP7
    Case 8
        Set objhttp = objHTTP8
    Case 9
        Set objhttp = objHTTP9
    Case 10
        Set objhttp = objHTTP10
    Case 11
        Set objhttp = objHTTP11
    Case 12
        Set objhttp = objHTTP12
    Case 13
        Set objhttp = objHTTP13
    Case 14
        Set objhttp = objHTTP14
    Case 15
        Set objhttp = objHTTP15
    Case 16
        Set objhttp = objHTTP16
    Case 17
        Set objhttp = objHTTP17
    Case 18
        Set objhttp = objHTTP18
    Case 19
        Set objhttp = objHTTP19
    End Select
End Sub


Sub eraseObjHTTPs()
' On Error Resume Next

    If IsArray(objHTTParr) And usingMacOSX = False And useQTforDataFetch = False Then
        For objHTTPnum = 1 To maxSimultaneousQueries
            Call assignObjHTTP
            If Not objhttp Is Nothing Then objhttp.abort
        Next objHTTPnum
    End If

    Set objhttp = Nothing
    Set objHTTP1 = Nothing
    Set objHTTP2 = Nothing
    Set objHTTP3 = Nothing
    Set objHTTP4 = Nothing
    Set objHTTP5 = Nothing
    Set objHTTP6 = Nothing
    Set objHTTP7 = Nothing
    Set objHTTP8 = Nothing
    Set objHTTP9 = Nothing
    Set objHTTP10 = Nothing
    Set objHTTP11 = Nothing
    Set objHTTP12 = Nothing
    Set objHTTP13 = Nothing
    Set objHTTP14 = Nothing
    Set objHTTP15 = Nothing
    Set objHTTP16 = Nothing
    Set objHTTP17 = Nothing
    Set objHTTP18 = Nothing
    Set objHTTP19 = Nothing

    If IsArray(queryArr) Then Erase queryArr
    If IsArray(objHTTParr) Then Erase objHTTParr
    Call deleteQueryTables
    queryTableResultStr = vbNullString
End Sub




Sub initializeObjHTTP()

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    If usingMacOSX = False And useQTforDataFetch = False Then


        Call setMSXML(objHTTP1)
        If useProxy = True Then objHTTP1.setProxy 2, proxyAddress
        If useProxyWithCredentials = True Then objHTTP1.setProxyCredentials proxyUsername, proxyPassword

        If queryCount >= 2 And maxSimultaneousQueries >= 2 Then
            Call setMSXML(objHTTP2)
            If useProxy = True Then objHTTP2.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP2.setProxyCredentials proxyUsername, proxyPassword
        End If

        If queryCount >= 3 And maxSimultaneousQueries >= 3 Then
            Call setMSXML(objHTTP3)
            If useProxy = True Then objHTTP3.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP3.setProxyCredentials proxyUsername, proxyPassword
        End If

        If queryCount >= 4 And maxSimultaneousQueries >= 4 Then
            Call setMSXML(objHTTP4)
            If useProxy = True Then objHTTP4.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP4.setProxyCredentials proxyUsername, proxyPassword
        End If
        If queryCount >= 5 And maxSimultaneousQueries >= 5 Then
            Call setMSXML(objHTTP5)
            If useProxy = True Then objHTTP5.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP5.setProxyCredentials proxyUsername, proxyPassword
        End If
        If queryCount >= 6 And maxSimultaneousQueries >= 6 Then
            Call setMSXML(objHTTP6)
            If useProxy = True Then objHTTP6.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP6.setProxyCredentials proxyUsername, proxyPassword
        End If
        If queryCount >= 7 And maxSimultaneousQueries >= 7 Then
            Call setMSXML(objHTTP7)
            If useProxy = True Then objHTTP7.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP7.setProxyCredentials proxyUsername, proxyPassword
        End If
        If queryCount >= 8 And maxSimultaneousQueries >= 8 Then
            Call setMSXML(objHTTP8)
            If useProxy = True Then objHTTP8.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP8.setProxyCredentials proxyUsername, proxyPassword
        End If
        If queryCount >= 9 And maxSimultaneousQueries >= 9 Then
            Call setMSXML(objHTTP9)
            If useProxy = True Then objHTTP9.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP9.setProxyCredentials proxyUsername, proxyPassword
        End If

        If queryCount >= 10 And maxSimultaneousQueries >= 10 Then
            Call setMSXML(objHTTP10)
            If useProxy = True Then objHTTP10.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP10.setProxyCredentials proxyUsername, proxyPassword
        End If

        If queryCount >= 11 And maxSimultaneousQueries >= 11 Then
            Call setMSXML(objHTTP11)
            If useProxy = True Then objHTTP11.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP11.setProxyCredentials proxyUsername, proxyPassword
        End If

        If queryCount >= 12 And maxSimultaneousQueries >= 12 Then
            Call setMSXML(objHTTP12)
            If useProxy = True Then objHTTP12.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP12.setProxyCredentials proxyUsername, proxyPassword
        End If

        If queryCount >= 13 And maxSimultaneousQueries >= 13 Then
            Call setMSXML(objHTTP13)
            If useProxy = True Then objHTTP13.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP13.setProxyCredentials proxyUsername, proxyPassword
        End If

        If queryCount >= 14 And maxSimultaneousQueries >= 14 Then
            Call setMSXML(objHTTP14)
            If useProxy = True Then objHTTP14.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP14.setProxyCredentials proxyUsername, proxyPassword
        End If
        If queryCount >= 15 And maxSimultaneousQueries >= 15 Then
            Call setMSXML(objHTTP15)
            If useProxy = True Then objHTTP15.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP15.setProxyCredentials proxyUsername, proxyPassword
        End If
        If queryCount >= 16 And maxSimultaneousQueries >= 16 Then
            Call setMSXML(objHTTP16)
            If useProxy = True Then objHTTP16.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP16.setProxyCredentials proxyUsername, proxyPassword
        End If
        If queryCount >= 17 And maxSimultaneousQueries >= 17 Then
            Call setMSXML(objHTTP17)
            If useProxy = True Then objHTTP17.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP17.setProxyCredentials proxyUsername, proxyPassword
        End If
        If queryCount >= 18 And maxSimultaneousQueries >= 18 Then
            Call setMSXML(objHTTP18)
            If useProxy = True Then objHTTP18.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP18.setProxyCredentials proxyUsername, proxyPassword
        End If
        If queryCount >= 19 And maxSimultaneousQueries >= 19 Then
            Call setMSXML(objHTTP19)
            If useProxy = True Then objHTTP19.setProxy 2, proxyAddress
            If useProxyWithCredentials = True Then objHTTP19.setProxyCredentials proxyUsername, proxyPassword
        End If

    End If
End Sub



