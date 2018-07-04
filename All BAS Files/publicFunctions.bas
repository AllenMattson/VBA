Attribute VB_Name = "publicFunctions"
Public Function getAccountData(authToken As String, Optional dataType As String = "profiles", Optional includeHeaders As Boolean = False) As Variant

'See documentation at http://www.automateanalytics.com/p/google-analytics-vba-functions.html

    On Error GoTo errhandler

    Dim requestStr As String
    Dim responseStr As String

    Dim URL As String

    Dim rivi As Long

    Dim resultArr As Variant
    Dim resultsArrCol1 As Variant
    Dim resultsArrCol2 As Variant
    Dim resultsArrCol3 As Variant
    Dim rowCount As Long

    Dim objhttp As Object
    Dim errorStr As String
    Dim errorCount As Long

    If authToken = "" Then
        ReDim resultArr(1 To 1, 1 To 1)
        resultArr(1, 1) = "Error: Authentication token missing, get one from AnalyticsFunctions.com"
        getAccountData = resultArr
        Exit Function
    End If

    If Left(authToken, 5) <> "pwfat" Then
        ReDim resultArr(1 To 1, 1 To 1)
        resultArr(1, 1) = "Error: Authentication token invalid, get one from AnalyticsFunctions.com"
        getAccountData = resultArr
        Exit Function
    End If



    URL = "https://supermetrics.com/api/getAccount?responseFormat=RSCL"

    requestStr = "token=" & uriEncode(authToken)
    requestStr = requestStr & "&datatype=" & dataType
    requestStr = requestStr & "&version=" & versionNumber
    requestStr = requestStr & "&rscL1=" & uriEncode(rscL1)
    requestStr = requestStr & "&rscL0=" & uriEncode(rscL0)
    requestStr = requestStr & "&rscL2=" & uriEncode(rscL2)
    requestStr = requestStr & "&rscL3=" & uriEncode(rscL3)
    requestStr = requestStr & "&rscL4=" & uriEncode(rscL4)
    If includeHeaders = False Then
        requestStr = requestStr & "&includeHeaders=0"
    Else
        requestStr = requestStr & "&includeHeaders=1"
    End If



runFetchAgain:


    Call setMSXML(objhttp)
    'objHTTP.setProxy 2, proxyAddress     'if connecting via proxy server, uncomment row and type proxy address
    'objHTTP.setProxyCredentials proxyUsername, proxyPassword     'if connecting via proxy server that requires authentication, uncomment row and type credentials
    objhttp.Open "POST", URL, True
    objhttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    objhttp.setTimeouts 1000000, 1000000, 1000000, 1000000
    objhttp.setOption 2, 13056
    On Error GoTo errHandler2
    objhttp.send (requestStr)
    On Error GoTo errhandler

    Do
        If objhttp.readyState = 4 Then Exit Do
        DoEvents
    Loop

    responseStr = objhttp.responsetext
    Set objhttp = Nothing

    Debug.Print Left(responseStr, 5000)

    If Left(responseStr, 7) <> "SUCCESS" Then
        ReDim resultArr(1 To 1, 1 To 1)
        resultArr(1, 1) = "Error: " & responseStr
        getAccountData = resultArr
        Exit Function
    End If



    responseStr = Right(responseStr, Len(responseStr) - 7)
    resultArr = Split(responseStr, rscL2)

    resultsArrCol1 = Split(resultArr(1), rscL3)   'goals: profid
    resultsArrCol2 = Split(resultArr(2), rscL3)   'goals: goalid
    If dataType <> "segments" Then resultsArrCol3 = Split(resultArr(3), rscL3)   'goals: goalname

    rowCount = UBound(resultsArrCol1) + 1

    If rowCount = 0 Then
        Debug.Print "No data found"
        ReDim resultArr(1 To 1, 1 To 1)
        resultArr(1, 1) = "Error: No data found"
        getAccountData = resultArr
        Exit Function
    End If

    ReDim resultArr(1 To rowCount, 1 To 3)

    For rivi = 1 To rowCount
        If dataType = "segments" Then
            resultArr(rivi, 1) = Left(resultsArrCol2(rivi - 1), 255)
            resultArr(rivi, 2) = Left(resultsArrCol1(rivi - 1), 255)
        Else
            resultArr(rivi, 1) = Left(resultsArrCol1(rivi - 1), 255)
            resultArr(rivi, 2) = Left(resultsArrCol2(rivi - 1), 255)
            resultArr(rivi, 3) = Left(resultsArrCol3(rivi - 1), 255)
        End If
    Next rivi

    resultArr = arrayReplace(resultArr, "%rscL1%", rscL1)
    resultArr = arrayReplace(resultArr, "%rscL2%", rscL2)
    resultArr = arrayReplace(resultArr, "%rscL3%", rscL3)

    getAccountData = resultArr

    Exit Function

errhandler:
    ReDim resultArr(1 To 1, 1 To 1)
    resultArr(1, 1) = "Error: " & Err.Description
    getAccountData = resultArr

    Exit Function

errHandler2:
    Resume runFetchAgain

End Function

Public Function getGAData(ByVal authToken As String, ByVal profiles As String, ByVal metrics As String, ByVal startDate As Date, ByVal endDate As Date, Optional ByVal filters As String, Optional ByVal dimensions As String, Optional ByVal segment As String, Optional sort As Variant = False, Optional includeHeaders As Boolean = False, Optional includeDimensionColumns As Variant = True, Optional maxResults As Long = 10000, Optional pivotDimensions As String = "", Optional settings As String = "", Optional maxCategories As Long = 20) As Variant
    getGAData = getData(authToken, profiles, metrics, startDate, endDate, filters, dimensions, segment, sort, includeHeaders, includeDimensionColumns, maxResults, pivotDimensions, settings, maxCategories)
End Function

Public Function getData(ByVal authToken As String, ByVal profiles As String, ByVal metrics As String, ByVal startDate As Date, ByVal endDate As Date, Optional ByVal filters As String, Optional ByVal dimensions As String, Optional ByVal segment As String, Optional sort As Variant = False, Optional includeHeaders As Boolean = False, Optional includeDimensionColumns As Variant = True, Optional maxResults As Long = 10000, Optional pivotDimensions As String = "", Optional settings As String = "", Optional maxCategories As Long = 20) As Variant
    If includeHeaders Then
        settings = settings & ",INCLUDE_HEADERS"
    Else
        settings = settings & ",NO_HEADERS"
    End If
    If Not includeDimensionColumns Then
        settings = settings & ",NO_DIMENSIONS"
    End If
    getData = Supermetrics(authToken, profiles, metrics, startDate, endDate, dimensions, pivotDimensions, filters, segment, sort, settings, maxResults, maxCategories)
End Function
Public Function Supermetrics(ByVal authToken As String, ByVal profiles As String, ByVal metrics As String, ByVal startDate As Date, ByVal endDate As Date, Optional ByVal dimensions As String, Optional ByVal pivotDimensions As String = "", Optional ByVal filters As String, Optional ByVal segment As String, Optional sort As Variant = False, Optional settings As String = "", Optional maxResults As Long = 10000, Optional maxCategories As Long = 20) As Variant

'See documentation at http://supermetrics.com/support/supermetrics-functions/

    On Error GoTo errhandler

    Dim startDateString As String
    Dim endDateString As String

    Dim URL As String

    Dim requestStr As String
    Dim resultArr As Variant

    If authToken = "" Then
        ReDim resultArr(1 To 1, 1 To 1)
        resultArr(1, 1) = "Error: Authentication token missing, get one from supermetrics.com/functions-login"
        Supermetrics = resultArr
        Exit Function
    End If

    If Left(authToken, 5) <> "pwfat" Then
        ReDim resultArr(1 To 1, 1 To 1)
        resultArr(1, 1) = "Error: Authentication token invalid, get one from supermetrics.com/functions-login"
        Supermetrics = resultArr
        Exit Function
    End If

    If startDate > endDate Then
        ReDim resultArr(1 To 1, 1 To 1)
        resultArr(1, 1) = "Error: Start date should be before end date"
        Supermetrics = resultArr
        Exit Function
    End If

    If maxResults = 0 Then maxResults = 10000

    dimensions = Replace(Replace(dimensions, ",", rscL1), "&", rscL1)
    If pivotDimensions <> vbNullString Then
        pivotDimensions = Replace(Replace(pivotDimensions, ",", rscL1), "&", rscL1)
        pivotDimensions = Replace(pivotDimensions, rscL1, "_pivot" & rscL1) & "_pivot"
        dimensions = dimensions & rscL1 & pivotDimensions
        If InStr(1, settings, "NO_HEADERS") = 0 And InStr(1, settings, "INCLUDE_HEADERS") = 0 Then settings = settings & ",INCLUDE_HEADERS"
    End If
    dimensions = uriEncode(dimensions)

    URL = "https://supermetrics.com/api/getData?responseFormat=RSCL"

    startDateString = Year(startDate) & "-" & Right("0" & Month(startDate), 2) & "-" & Right("0" & Day(startDate), 2)
    endDateString = Year(endDate) & "-" & Right("0" & Month(endDate), 2) & "-" & Right("0" & Day(endDate), 2)

    requestStr = requestStr & "token=" & uriEncode(authToken)
    requestStr = requestStr & "&system=Excel"
    requestStr = requestStr & "&version=" & uriEncode(versionNumber)
    If separatorList = vbNullString Then
        separatorList = "&rscL1=" & uriEncode(rscL1)
        separatorList = separatorList & "&rscL0=" & uriEncode(rscL0)
        separatorList = separatorList & "&rscL2=" & uriEncode(rscL2)
        separatorList = separatorList & "&rscL3=" & uriEncode(rscL3)
        separatorList = separatorList & "&rscL4=" & uriEncode(rscL4)
    End If
    requestStr = requestStr & separatorList

    requestStr = requestStr & "&profiles=" & uriEncode(Replace(Replace(profiles, ",", rscL1), "&", rscL1))
    requestStr = requestStr & "&metrics=" & uriEncode(Replace(Replace(metrics, ",", rscL1), "&", rscL1))
    requestStr = requestStr & "&dimensions=" & dimensions
    requestStr = requestStr & "&filterStr=" & uriEncode(filters)
    requestStr = requestStr & "&segment=" & uriEncode(Replace(segment, rscL2, "%rscL2%"))
    requestStr = requestStr & "&dateRangeType=fixed"
    requestStr = requestStr & "&start-date=" & uriEncode(startDateString)
    requestStr = requestStr & "&end-date=" & uriEncode(endDateString)

    If sort = True Then
        requestStr = requestStr & "&sort=metricdesc"
    ElseIf sort = False Then
        requestStr = requestStr & "&sort=dimensionasc"
    Else
        requestStr = requestStr & "&sort=" & sort
    End If


    requestStr = requestStr & "&maxResultsAll=" & maxResults
    requestStr = requestStr & "&maxCategories=" & maxCategories
    requestStr = requestStr & "&settings=" & uriEncode(settings)

    Debug.Print "Request: " & requestStr

    Dim objhttp As Object

runFetchAgain:

    Call setMSXML(objhttp)

    'objHTTP.setProxy 2, proxyAddress     'if connecting via proxy server, uncomment row and type proxy address
    'objHTTP.setProxyCredentials proxyUsername, proxyPassword     'if connecting via proxy server that requires authentication, uncomment row and type credentials
    objhttp.Open "POST", URL, False
    objhttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
    objhttp.setTimeouts 1000000, 1000000, 1000000, 1000000
    objhttp.setOption 2, 13056
    On Error GoTo errHandler2
    objhttp.send (requestStr)
    On Error GoTo errhandler

    Dim gaResponse As String
    gaResponse = objhttp.responsetext

    Debug.Print "GA data feed response: " & Left(gaResponse, 3000)


    resultArr = parseResponseToArr(gaResponse)
    Supermetrics = resultArr

    Exit Function


errhandler:

    ReDim TempArray(1 To 1, 1 To 1)
    TempArray(1, 1) = "Error: Fetching data failed: " & Err.Description
    Supermetrics = TempArray

    Exit Function

errHandler2:
    Resume runFetchAgain

End Function


Public Function parseResponseToArr(responseStr As Variant) As Variant

    On Error Resume Next

    Dim TempArray As Variant
    Dim TempArray2 As Variant
    Dim dataArr As Variant
    Dim dataArrJ As Variant

    Dim rowNum As Long
    Dim rowCount As Long
    Dim colNum As Long
    Dim colCount As Long
    Dim noteStr As String
    Dim headersArr As Variant
    Dim licenseResponseStr As String
    Dim convertRSCL As Boolean
    Dim numericColumnsStart As Integer
    Dim numericRowsStart As Integer
    Dim i As Integer

    responseStr = Replace(responseStr, vbCrLf, "")
    responseStr = Replace(responseStr, vbLf, "")
    responseStr = Replace(responseStr, vbCr, "")

    noteStr = Split(responseStr, rscL1)(0)

    numericColumnsStart = CInt(parseVarFromStr(noteStr, "NUMERIC_FORMAT_COLUMNS_START", rscL2))
    numericRowsStart = CInt(parseVarFromStr(noteStr, "NUMERIC_FORMAT_ROWS_START", rscL2))

    If parseVarFromStr(noteStr, "CONVERT_RSCL", rscL2) = "TRUE" Then
        convertRSCL = True
    Else
        convertRSCL = False
    End If

    If parseVarFromStr(noteStr, "STATUS", rscL2) = "ERROR" Then
        ReDim TempArray(1 To 1, 1 To 1)
        TempArray(1, 1) = "Error: " & parseVarFromStr(responseStr, "ERROR", rscL2)
        parseResponseToArr = TempArray
        Exit Function
    End If


    headersArr = Split(Split(responseStr, rscL1)(1), rscL2)
    dataArr = Split(Split(responseStr, rscL1)(2), rscL2)
    licenseResponseStr = Split(responseStr, rscL1)(3)

    colCount = UBound(headersArr)

    rowCount = UBound(dataArr)


    If rowCount < 0 Or colCount < 0 Then
        ReDim TempArray(1 To 1, 1 To 1)
        TempArray(1, 1) = "Error: No data found"
        parseResponseToArr = TempArray
        Exit Function
    End If


    ReDim TempArray(1 To rowCount + 1, 1 To colCount + 1)
    ReDim dataArrJ(0 To rowCount)

    For rowNum = 0 To rowCount
        dataArrJ(rowNum) = Split(dataArr(rowNum), rscL3)
    Next rowNum


    For rowNum = 0 To rowCount
        For colNum = 0 To colCount
            If rowNum >= numericRowsStart And colNum >= numericColumnsStart Then
                If IsNumeric(val(dataArrJ(rowNum)(colNum))) Then
                    TempArray(rowNum + 1, colNum + 1) = val(dataArrJ(rowNum)(colNum))
                Else
                    TempArray(rowNum + 1, colNum + 1) = Left(dataArrJ(rowNum)(colNum), 255)
                End If
            Else
                TempArray(rowNum + 1, colNum + 1) = Left(replaceRSCL(dataArrJ(rowNum)(colNum)), 255)
            End If
        Next colNum
        If i > 200 Then
            DoEvents
            i = 0
        End If
    Next rowNum

    If convertRSCL Then TempArray = arrayReplaceRSCL(TempArray, numericColumnsStart)

    parseResponseToArr = TempArray

End Function

Public Function getTweets(token As String, searchTerm As String, Optional columnsToFetch As String = "time,twitter_name,tweet", Optional maxResults As Integer = 100, Optional includeHeaders As Boolean = True, Optional resultType As String = "recent", Optional languageCode As String, Optional geoCode As String, Optional untilDate As String, Optional locale As String, Optional timeZone As String = "UTC", Optional fromMacro = False) As Variant

    Dim URL As String
    Dim requestStr As String
    Dim responseArr As Variant

    Dim objhttp As Object



    URL = "https://supermetrics.com/api/getData?responseFormat=RSCL"

    requestStr = "q=" & uriEncode(searchTerm)
    requestStr = requestStr & "&token=" & uriEncode(token)
    requestStr = requestStr & "&dataSource=TW"
    requestStr = requestStr & "&format=RSCL"
    requestStr = requestStr & "&columns=" & uriEncode(columnsToFetch)
    requestStr = requestStr & "&maxResults=" & maxResults
    requestStr = requestStr & "&result_type=" & resultType
    requestStr = requestStr & "&timezone=" & uriEncode(timeZone)

    If separatorList = vbNullString Then
        separatorList = "&rscL1=" & uriEncode(rscL1)
        separatorList = separatorList & "&rscL0=" & uriEncode(rscL0)
        separatorList = separatorList & "&rscL2=" & uriEncode(rscL2)
        separatorList = separatorList & "&rscL3=" & uriEncode(rscL3)
        separatorList = separatorList & "&rscL4=" & uriEncode(rscL4)
    End If
    requestStr = requestStr & separatorList


    If languageCode <> "" Then
        requestStr = requestStr & "&lang=" & languageCode
    End If
    If locale <> "" Then
        requestStr = requestStr & "&locale=" & locale
    End If
    If untilDate <> "" Then
        requestStr = requestStr & "&until=" & uriEncode(untilDate)
    End If
    If geoCode <> "" Then
        requestStr = requestStr & "&geocode=" & uriEncode(geoCode)
    End If
    If includeHeaders Then
        requestStr = requestStr & "&includeHeaders=1"
    Else
        requestStr = requestStr & "&includeHeaders=0"
    End If

runFetchAgain:
    If useQTforDataFetch Then
        requestStr = requestStr & "&chrEncode=1"
        Call fetchDataWithQueryTableDirect(URL, requestStr, True, True)
        responseStr = queryTableResultStr
    Else
        Call getProxySettingsIfNeeded
        Call setMSXML(objhttp)

        If useProxy = True Then objhttp.setProxy 2, proxyAddress
        If useProxyWithCredentials = True Then objhttp.setProxyCredentials proxyUsername, proxyPassword
        objhttp.Open "POST", URL, True
        objhttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        objhttp.setTimeouts 1000000, 1000000, 1000000, 1000000
        objhttp.setOption 2, 13056
        On Error GoTo errHandler2
        objhttp.send (requestStr)
        On Error GoTo 0

        Do
            objhttp.waitForResponse 0
            If objhttp.readyState = 4 Then Exit Do
            If fromMacro Then Call updateProgressIterationBoxes
        Loop


        responseStr = objhttp.responsetext
    End If
    If debugMode Then
        Debug.Print "req: " & requestStr
        Debug.Print "response: " & Left(responseStr, 3000)
    End If
    responseArr = parseResponseToArr(responseStr)

    getTweets = responseArr

    Exit Function

errHandler2:
    Resume runFetchAgain

End Function




