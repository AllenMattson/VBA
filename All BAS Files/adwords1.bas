Attribute VB_Name = "adwords1"
Option Private Module
Option Explicit


Public Function getACtoken(emailLoc As String, passwordLoc As String) As String

    Dim postStr As String
    Dim URL As String
    Dim objhttp As Object
    postStr = "email=" & uriEncode(encrypt(emailLoc))
    postStr = postStr & "&pw=" & uriEncode(encrypt(passwordLoc))

    URL = "https://supermetrics.com/api/getACtoken?responseFormat=RSCL"

    If usingMacOSX = True Or useQTforDataFetch = True Then
        Call fetchDataWithQueryTableDirect(URL, postStr)
        authToken = queryTableResultStr
    Else
        Call setMSXML(objhttp)
        If useProxy = True Then objhttp.setProxy 2, proxyAddress
        objhttp.Open "POST", URL, False
        If useProxyWithCredentials = True Then objhttp.setProxyCredentials proxyUsername, proxyPassword
        objhttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        objhttp.setTimeouts 1000000, 1000000, 1000000, 1000000
        objhttp.setOption 2, 13056
        objhttp.send (postStr)
        authToken = objhttp.responsetext
        Set objhttp = Nothing
    End If
    getACtoken = authToken
End Function

Sub checkSelectedFields(Optional userInitiated As Boolean = False)
    On Error Resume Next
    Call checkOperatingSystem
    If usingMacOSX And Not userInitiated Then Exit Sub

    ' Debug.Print "Running checkFields..."
    Dim firstField As Boolean
    Dim fieldsArr As Variant
    Dim i As Long
    Dim dimensions As String
    Dim metrics As String
    Call setDatasourceVariables

    firstField = True
    fieldsArr = Range("dimension1name" & varsuffix).Offset(, 1).Resize(14, 1).value
    For i = 1 To UBound(fieldsArr)
        If fieldNameIsOk(fieldsArr(i, 1)) Then
            If firstField = False Then dimensions = dimensions & "#"
            dimensions = dimensions & fieldsArr(i, 1)
            firstField = False
        End If
    Next i

    dimensions = Replace(dimensions, rscL1, "#")

    firstField = True
    fieldsArr = Range("metric1name" & varsuffix).Offset(, 2).Resize(13, 1).value
    For i = 1 To UBound(fieldsArr)
        If fieldNameIsOk(fieldsArr(i, 1)) Then
            If firstField = False Then metrics = metrics & "#"
            metrics = metrics & fieldsArr(i, 1)
            firstField = False
        End If
    Next i
    metrics = Replace(metrics, rscL1, "#")

    Dim fieldsAreOK As Boolean
    fieldsAreOK = False
    If metrics = vbNullString And dimensions = vbNullString Then
        fieldsAreOK = True
    ElseIf dimensions = vbNullString And dataSource = "GA" Then
        fieldsAreOK = True
    ElseIf dimensions = vbNullString And InStr(1, metrics, "#") = 0 Then
        fieldsAreOK = True
    ElseIf metrics = vbNullString And InStr(1, dimensions, "#") = 0 Then
        fieldsAreOK = True
    ElseIf dataSource = "GA" And InStr(1, dimensions, "mcf:") = 0 And InStr(1, metrics, "mcf:") = 0 Then
        fieldsAreOK = True
    ElseIf checkFieldStatus(metrics, dimensions) Then
        fieldsAreOK = True
    Else
        fieldsAreOK = False
    End If


    If fieldsAreOK Then
        If userInitiated Or configsheet.Shapes("illegalFieldsWarning" & varsuffix).Visible = True Then
            Call unprotectSheets
            configsheet.Shapes("illegalFieldsWarning" & varsuffix).Visible = False
            If userInitiated Then configsheet.Shapes("fieldsOKnote").Visible = True
            Call protectSheets
        End If
    Else
        If userInitiated Or configsheet.Shapes("illegalFieldsWarning" & varsuffix).Visible = False Then
            Call unprotectSheets
            configsheet.Shapes("illegalFieldsWarning" & varsuffix).Visible = True
            If userInitiated Then configsheet.Shapes("fieldsOKnote").Visible = False
            Call protectSheets
        End If
    End If

End Sub







Public Function checkFieldStatus(metrics As String, Optional dimensions As String)

    On Error Resume Next
    '  If debugMode = True Then On Error GoTo 0

    checkFieldStatus = True
    Dim responseStr As String

    Dim postStr As String
    Dim password As String
    Dim objHTTPACfields As Object
    Dim URL As String

    postStr = "metrics=" & metrics
    postStr = postStr & "&dimensions=" & dimensions
    postStr = postStr & "&appid=" & appID
    postStr = postStr & "&dataSource=" & dataSource

    URL = "https://supermetrics.com/api/checkFields?responseFormat=RSCL"


runFetchAgain:
    If usingMacOSX = True Or useQTforDataFetch = True Then
        Call fetchDataWithQueryTableDirect(URL, postStr)
        responseStr = queryTableResultStr
    Else
        Call setMSXML(objHTTPACfields)
        If useProxy = True Then objHTTPACfields.setProxy 2, proxyAddress
        objHTTPACfields.Open "POST", URL, False
        If useProxyWithCredentials = True Then objHTTPACfields.setProxyCredentials proxyUsername, proxyPassword
        objHTTPACfields.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        objHTTPACfields.setTimeouts 1000000, 1000000, 1000000, 1000000
        objHTTPACfields.setOption 2, 13056

        On Error GoTo errhandler

        objHTTPACfields.send (postStr)
        If debugMode Then Debug.Print objHTTPACfields.responsetext
        responseStr = objHTTPACfields.responsetext
        On Error Resume Next
    End If
    If InStr(1, responseStr, "FIELDSOK") > 0 Then
        checkFieldStatus = True
    Else
        checkFieldStatus = False
    End If


    Exit Function

errhandler:
    Resume runFetchAgain

End Function



Public Function encodeStr(str As String)
    On Error Resume Next

    Dim Counter As Long

    For Counter = 1 To Len(str)
        If Asc(Mid(str, Counter, 1)) > 65 And Asc(Mid(str, Counter, 1)) < 88 Then
            Mid(str, Counter, 1) = Chr(Asc(Mid(str, Counter, 1)) + 2)
        ElseIf Asc(Mid(str, Counter, 1)) > 100 And Asc(Mid(str, Counter, 1)) < 122 Then
            Mid(str, Counter, 1) = Chr(Asc(Mid(str, Counter, 1)) - 2)
        End If
    Next

    encodeStr = str
End Function



Public Function copyArray(arr As Variant, lbound1 As Long, ubound1 As Long, lbound2 As Long, ubound2 As Long)

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim rivi As Long
    Dim sar As Long

    Dim tempArr() As Variant
    ReDim tempArr(lbound1 To ubound1, lbound2 To ubound2)

    For rivi = lbound1 To ubound1
        For sar = lbound2 To ubound2
            If rivi >= LBound(arr, 1) And rivi <= UBound(arr, 1) And sar >= LBound(arr, 2) And sar <= UBound(arr, 2) Then
                tempArr(rivi, sar) = arr(rivi, sar)
            End If
        Next sar
    Next rivi

    copyArray = tempArr

End Function


Public Function changeArrayDataTypes(arr As Variant, Optional firstMetricCol As Long = 0)

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    aika = Timer

    Dim changeToNumericArr() As Variant
    ReDim changeToNumericArr(1 To UBound(arr, 2))
    Dim sar As Long
    Dim rivi As Long

    If firstMetricCol = 0 Then
        For sar = 1 To UBound(arr, 2)
            If isMetric(arr(0, sar)) = True Then
                For rivi = 1 To UBound(arr, 1)
                    If arr(rivi, sar) <> "" Then
                        arr(rivi, sar) = val(Replace(arr(rivi, sar), ",", ""))
                    End If
                Next rivi
            End If
        Next sar
    Else
        If firstMetricCol > UBound(arr, 2) Then
            changeArrayDataTypes = arr
            Debug.Print "changeArrayDataTypes firstMetricCol larger than column count " & firstMetricCol & " vs " & UBound(arr, 2)
            Exit Function
        End If
        For sar = firstMetricCol To UBound(arr, 2)
            For rivi = 1 To UBound(arr, 1)
                If arr(rivi, sar) <> "" Then
                    If Right(arr(rivi, sar), 1) = "%" Then
                        arr(rivi, sar) = val(Replace(Replace(arr(rivi, sar), ",", ""), "%", "")) / 100
                    Else
                        arr(rivi, sar) = val(Replace(arr(rivi, sar), ",", ""))
                    End If
                End If
            Next rivi
        Next sar
    End If

    changeArrayDataTypes = arr

    Debug.Print "AW change data types time: " & Timer - aika

End Function




Public Function compressArrayToTopValues(ByVal arr As Variant, Optional filterColumn As Long = 1, Optional numberOfResults As Long = 10)

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim resultArr() As Variant
    Dim usedRowsArr() As Variant
    Dim rivi As Long
    Dim sar As Long
    Dim arvo As Variant
    Dim valuesFound As Long
    Dim uniqueValuesFound As Long

    Dim rankNum As Long
    Dim largestValuesArr As Variant

    Dim rowCount As Long
    rowCount = UBound(arr)


    Dim filterColumn2 As Long   'use the filterColumn parameter first, then iterate through other columns

    valuesFound = 0
    uniqueValuesFound = 0
    If numberOfResults >= UBound(arr, 1) Then
        resultArr = arr
        numberOfResults = UBound(arr, 1)
    Else
        ReDim resultArr(1 To numberOfResults, LBound(arr, 2) To UBound(arr, 2))
        ReDim largestValuesArr(1 To numberOfResults)
        ReDim usedRowsArr(LBound(arr) To rowCount)

        For filterColumn2 = filterColumn To UBound(arr, 2)
            For rankNum = 1 To numberOfResults
                arvo = Application.Large(Application.Index(arr, , filterColumn2), rankNum)
                If Not Application.IsError(arvo) Then
                    uniqueValuesFound = uniqueValuesFound + 1
                    For rivi = 1 To rowCount
                        If usedRowsArr(rivi) = False Then
                            If arr(rivi, filterColumn2) = arvo Then
                                valuesFound = valuesFound + 1
                                usedRowsArr(rivi) = True
                                For sar = LBound(arr, 2) To UBound(arr, 2)
                                    resultArr(valuesFound, sar) = arr(rivi, sar)
                                    If valuesFound = numberOfResults Then GoTo exitLoops
                                Next sar
                            End If
                        End If
                    Next rivi
                    '  largestValuesArr(uniqueValuesFound) = Application.Large(Application.Index(arr, , filterColumn), rankNum)
                Else
                    Exit For
                End If
            Next rankNum
        Next filterColumn2

exitLoops:


    End If

    compressArrayToTopValues = resultArr

End Function





Public Function removeEmptyRowsFromEndOfArray(ByVal arr As Variant)

'tarkastaa vain ekan sarakkeen

    Dim rivi As Long
    Dim sar As Long
    Dim resultArr() As Variant
    Dim exitLoop As Boolean


    Dim lastRowOfData As Long

    Debug.Print "AW Rows in resultArr before truncate: " & UBound(arr, 1)


    lastRowOfData = UBound(arr, 1)
    For rivi = 1 To UBound(arr, 1)
        If arr(rivi, 1) = vbNullString Then
            exitLoop = True
            For sar = LBound(arr, 2) To UBound(arr, 2)
                If arr(rivi, sar) <> vbNullString Then
                    exitLoop = False
                    Exit For
                End If
            Next sar

            If exitLoop = True Then
                lastRowOfData = rivi - 1
                Exit For
            End If

        End If
    Next rivi


    If lastRowOfData = UBound(arr, 1) Then
        Debug.Print "AW Rows in resultArr after truncate: " & UBound(arr, 1)
        removeEmptyRowsFromEndOfArray = arr
        Exit Function
    Else
        ReDim resultArr(LBound(arr, 1) To lastRowOfData, LBound(arr, 2) To UBound(arr, 2))
        For rivi = LBound(resultArr, 1) To UBound(resultArr, 1)
            For sar = LBound(resultArr, 2) To UBound(resultArr, 2)
                resultArr(rivi, sar) = arr(rivi, sar)
                '      Debug.Print "arr(rivi, sar) " & arr(rivi, sar)
            Next sar
        Next rivi
    End If

    Debug.Print "AW Rows in resultArr after truncate: " & UBound(resultArr, 1)

    removeEmptyRowsFromEndOfArray = resultArr
    '  Debug.Print "removeEmptyRowsFromEndOfArray ended"
End Function

Public Function isTimeDimension(ByVal fieldName) As Boolean
    fieldName = LCase(fieldName)
    isTimeDimension = False
    If fieldName = "hour" Or fieldName = "day" Or fieldName = "date" Or fieldName = "gregoriandate" Or fieldName = "dayofweek" Or fieldName = "dayofmonth" Or fieldName = "weekday" Or fieldName = "week" Or fieldName = "weekus" Or fieldName = "weekiso" Or fieldName = "month" Or fieldName = "quarter" Or fieldName = "year" Then isTimeDimension = True
End Function

Public Function isMetric(fieldName) As Boolean

    Call setDatasourceVariables

    isMetric = False

    If Application.CountIf(Range("metrics" & varsuffix).Offset(0, 2).Resize(10000, 1), fieldName) > 0 Then
        isMetric = True  'check request codes
    ElseIf Application.CountIf(Range("metrics" & varsuffix).Offset(0, 9).Resize(10000, 1), fieldName) > 0 Then
        isMetric = True  'check response codes
    ElseIf Application.CountIf(Range("metrics" & varsuffix).Offset(0, 1).Resize(10000, 1), fieldName) > 0 Then
        isMetric = True  'check disp names
    Else
        If LCase$(fieldName) = "positionssum" Then isMetric = True
        If LCase$(fieldName) = "avgposition" Then isMetric = True
        If LCase$(fieldName) = "avgerageposition" Then isMetric = True
    End If

    'Calculated metrics used within the AW functions

    Debug.Print "Is " & dataSource & " metric " & fieldName & ": " & isMetric

End Function



Function VBAWeekNum(d As Date, Optional FWDayArg As Integer = 1) As Integer
' You can see examples of this at
' http://www.cpearson.com/excel/weeknum.htm
    VBAWeekNum = CInt(Format(d, "ww", FWDayArg))
End Function

