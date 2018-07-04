Attribute VB_Name = "OAuth"
Option Private Module
Option Explicit

Dim randStr As String

Sub testConnectionToSupermetrics()

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim authResponse As String
    Dim objHTTPauth As Object
    Dim URL As String
    Dim errorCount As Long
    errorCount = 0

    URL = "https://supermetrics.com/api/testConnection?responseFormat=RSCL"


tryAgain:
    If usingMacOSX = True Or useQTforDataFetch = True Then
        Call fetchDataWithQueryTableDirect(URL, "")
        authResponse = queryTableResultStr
    Else
        Call setMSXML(objHTTPauth)
        If useProxy = True Then objHTTPauth.setProxy 2, proxyAddress
        objHTTPauth.Open "GET", URL, False
        If useProxyWithCredentials = True Then objHTTPauth.setProxyCredentials proxyUsername, proxyPassword
        objHTTPauth.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        objHTTPauth.setTimeouts 100000, 100000, 100000, 100000
        objHTTPauth.setOption 2, 13056
        On Error GoTo connErr
        objHTTPauth.send ("")
        On Error Resume Next
        If debugMode = True Then On Error GoTo 0

        authResponse = objHTTPauth.responsetext

        Set objHTTPauth = Nothing
    End If
    Debug.Print "Connecting with Supermetrics: " & authResponse

    If authResponse = "Connection OK" Then Exit Sub

    MsgBox "Connecting with Supermetrics servers (supermetrics.com) failed. A firewall in your machine or network may be blocking the connection. This may also be a temporary problem - please try again later. It is possible that Supermetrics servers are currently down; if this is the case, it is usually announced at twitter.com/Supermetrics.", , "Failed to connect with Supermetrics"

    Exit Sub
connErr:
    errorCount = errorCount + 1
    If errorCount < 3 Then
        Resume tryAgain
    Else
        MsgBox "Connecting with Supermetrics servers (supermetrics.com) failed. A firewall in your machine or network may be blocking the connection. This may also be a temporary problem - please try again later. It is possible that Supermetrics servers are currently down; if this is the case, it is usually announced at twitter.com/Supermetrics.", , "Failed to connect with Supermetrics"
    End If

End Sub


Sub openOAuthAuthorizationPage()

    On Error GoTo errhandler
    If debugMode = True Then On Error GoTo 0

    Call testConnectionToSupermetrics

    Dim URL As String
    Dim shortURL As String
    Dim errorCount As Integer
    errorCount = 0
    profilesStr = ""
    Dim emailTimer As Double


    Randomize
    randStr = "r" & Year(Date) & Month(Date) & Day(Date) & genRandomString(25)
    If debugMode = True Then Debug.Print "randStr: " & randStr


    URL = "https://supermetrics.com/login/?r=" & randStr & "&appid=" & appID & "&version=" & versionNumber & "&service=" & dataSource & "&rid=" & randID

    If usingMacOSX = True Then
        URL = URL & "&os=mac"
    Else
        URL = URL & "&os=win"
    End If

    ' URL = URL & "&system=" & uriEncode(OSandExcelVersion)

    If dataSource = "AC" Then
        URL = URL & "&OAuthAC=1"
    End If


    On Error GoTo nonSSLURL
    ActiveWorkbook.FollowHyperlink Address:=URL
    On Error GoTo errhandler


    Application.Wait (Now + TimeValue("00:00:01"))

    If usingMacOSX = False Then ProgressBox.Show False



    progresspct = 7
    If loginType = "SECONDARY" Then
        Call updateProgress(progresspct, "Supermetrics Data Grabber is opening an authorization page in your Internet browser. Please log in with the user you wish to add, and approve the authorization to continue.")
    Else
        Call updateProgress(progresspct, "Supermetrics Data Grabber is opening an authorization page in your Internet browser. Please read that page and approve the authorization to continue.")
    End If

    shortURL = shortenURL(URL & "&linktype=tu")

    If shortURL <> vbNullString Then
        If loginType = "SECONDARY" Then
            Call updateProgress(progresspct, "Supermetrics Data Grabber is opening an authorization page in your Internet browser. Please log in with the user you wish to add, and approve the authorization to continue.", "If the authorization page has not been automatically opened in your browser, you can find it from this address: " & shortURL)
        Else
            Call updateProgress(progresspct, "Supermetrics Data Grabber is opening an authorization page in your Internet browser. Please read that page and approve the authorization to continue.", "If the authorization page has not been automatically opened in your browser, you can find it from this address: " & shortURL)
        End If
    End If

    If usingMacOSX = False Then ProgressBox.stopButton.Visible = True


    email = getEmail()



    If Left(email, 6) = "Error:" Then
        Call hideProgressBox
        MsgBox "An error occurred when trying to add your login. Please try again. The error message is: " & email
        End
    End If



    Exit Sub


errhandler:

    stParam2 = "OPENOAUTHPAGEERROR|" & Err.Number & "|" & Err.Description & "|" & proxyAddress
    Call checkE(email, dataSource, , True)
    Resume Next


nonSSLURL:
    URL = Replace(URL, "https://", "http://")
    Resume

End Sub









Public Function getEmail() As String

    On Error GoTo errhandler
    If debugMode = True Then On Error GoTo 0

    Dim errorCount As Integer
    Dim objHTTPemail As Object
    Dim objHTTPemail2 As Object
    Dim authResponse As String
    Dim URL As String
    Dim URL2 As String
    Dim requestStr As String
    Dim emailResponse As String
    Dim errorStr As String

    Dim emailFound As Boolean
    emailFound = False
    errorCount = 0

    segmentsStr = ""
    goalsStr = ""
    profilesStr = ""

    URL = "https://supermetrics.com/api/getAuthAndAccount?responseFormat=RSCL"
    URL2 = "https://supermetrics.com/api/getAuth?responseFormat=RSCL"


    requestStr = "randnum=" & randStr
    requestStr = requestStr & "&service=" & dataSource
    requestStr = requestStr & "&system=" & uriEncode(OSandExcelVersion)
    requestStr = requestStr & "&appid=" & appID & "&version=" & versionNumber & "&rid=" & randID
    If usingMacOSX = True Then
        requestStr = requestStr & "&chrencode=true"
        requestStr = requestStr & "&urlencode=true"
    End If
    requestStr = requestStr & "&encoding=light"
    requestStr = requestStr & "&arrayType=combined2"

    If separatorList = vbNullString Then
        separatorList = "&rscL1=" & uriEncode(rscL1)
        separatorList = separatorList & "&rscL2=" & uriEncode(rscL2)
        separatorList = separatorList & "&rscL3=" & uriEncode(rscL3)
        separatorList = separatorList & "&rscL4=" & uriEncode(rscL4)
    End If

    requestStr = requestStr & separatorList

fetchAgain:
    If usingMacOSX = True Or useQTforDataFetch = True Then
        Call fetchDataWithQueryTableDirect(URL, requestStr, True, True)
        authResponse = queryTableResultStr
    Else
        Call setMSXML(objHTTPemail)
        If useProxy = True Then objHTTPemail.setProxy 2, proxyAddress
        objHTTPemail.Open "POST", URL, True
        If useProxyWithCredentials = True Then objHTTPemail.setProxyCredentials proxyUsername, proxyPassword
        objHTTPemail.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        objHTTPemail.setTimeouts 1000000, 1000000, 1000000, 1000000
        objHTTPemail.setOption 2, 13056
        objHTTPemail.send (requestStr)

        Call setMSXML(objHTTPemail2)
        If useProxy = True Then objHTTPemail2.setProxy 2, proxyAddress
        objHTTPemail2.Open "POST", URL2, True
        If useProxyWithCredentials = True Then objHTTPemail2.setProxyCredentials proxyUsername, proxyPassword
        objHTTPemail2.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        objHTTPemail2.setTimeouts 1000000, 1000000, 1000000, 1000000
        objHTTPemail2.setOption 2, 13056
        objHTTPemail2.send (requestStr)

        Do
            objHTTPemail.waitForResponse 0
            If Not objHTTPemail2 Is Nothing Then objHTTPemail2.waitForResponse 0
            If objHTTPemail.readyState = 4 Then Exit Do
            If emailFound = False Then
                If objHTTPemail2.readyState = 4 Then
                    emailResponse = objHTTPemail2.responsetext
                    usernameDisp = parseVarFromStr(emailResponse, "USERNAMEDISP", rscL1)
                    email = parseVarFromStr(emailResponse, "EMAIL", rscL1)
                    If email <> vbNullString And InStr(1, LCase(email), "error") = 0 Then
                        If usernameDisp = "" Then usernameDisp = email
                        Call updateProgress(19, "Fetching account data...", "Authentication successful for " & usernameDisp)
                    Else
                        Call updateProgress(18, "Fetching account data...")
                    End If
                    emailFound = True
                    Set objHTTPemail2 = Nothing
                End If
            End If
            Call updateProgressIterationBoxes
        Loop

        If debugMode = True Then Debug.Print requestStr
        authResponse = objHTTPemail.responsetext
        If debugMode = True Then Debug.Print authResponse
        Set objHTTPemail = Nothing
    End If


    If authResponse = vbNullString Then authResponse = emailResponse

    email = parseVarFromStr(authResponse, "EMAIL", rscL1)
    usernameDisp = parseVarFromStr(authResponse, "USERNAMEDISP", rscL1)
    authToken = parseVarFromStr(authResponse, "TOKEN", rscL1)
    If dataSource = "GA" Or dataSource = "AW" Or dataSource = "YT" Then
        profilesStr = parseVarFromStr(authResponse, "PROFILES", rscL1)
        If dataSource = "GA" Then
            goalsStr = parseVarFromStr(authResponse, "GOALS", rscL1)
            segmentsStr = parseVarFromStr(authResponse, "SEGMENTS", rscL1)
        End If
    End If


    errorStr = parseVarFromStr(authResponse, "ERROR", rscL1)
    If errorStr = vbNullString Then errorStr = parseVarFromStr(authResponse, "ERROR", "|")
    If errorStr = vbNullString Then errorStr = parseVarFromStr(authResponse, "ERROR", "%")
    If errorStr = vbNullString And (email = vbNullString Or authToken = vbNullString) Then errorStr = "Authentication error"

    If errorStr <> vbNullString Then
        Call hideProgressBox
        Call protectSheets
        MsgBox "An error occurred when trying to add your login. Please try again. The error message is: " & errorStr
        End
    End If



    If usernameDisp = vbNullString Then usernameDisp = email

    getEmail = email
    Exit Function


    Exit Function

errhandler:


    stParam2 = "OAUTHEMAILERROR|" & Err.Number & "|" & Err.Description
    Call checkE(email, dataSource, True)
    Resume Next

End Function

Public Function refreshToken(ByVal oldToken) As String

    On Error GoTo errhandler
    '  If debugMode = True Then On Error GoTo 0

    Dim errorCount As Integer
    Dim objHTTPemail As Object
    Dim authResponse As String
    Dim URL As String
    Dim requestStr As String

    If debugMode Then Debug.Print "TOKEN REFRESH, old token: " & oldToken

    If dataSource = "AC" Then
        refreshToken = oldToken
        Exit Function
        '    ElseIf dataSource = "GW" Then
        '        refreshToken = getCLtoken(Sheets("cred").Cells(16, 1).value, decrypt(Sheets("cred").Cells(17, 1).value))
        '        Range("authtokenGW").value = refreshToken
        '        Exit Function
    End If

    errorCount = 0

    URL = "https://supermetrics.com/api/refreshToken?responseFormat=RSCL"

    requestStr = "token=" & oldToken
    requestStr = requestStr & "&appid=" & appID & "&version=" & versionNumber & "&rid=" & randID
    requestStr = requestStr & "&system=" & uriEncode(OSandExcelVersion)

fetchAgain:
    If usingMacOSX = True Or useQTforDataFetch = True Then
        Call fetchDataWithQueryTableDirect(URL, requestStr)
        authResponse = queryTableResultStr
    Else
        Call setMSXML(objHTTPemail)
        If useProxy = True Then objHTTPemail.setProxy 2, proxyAddress
        objHTTPemail.Open "POST", URL, True
        If useProxyWithCredentials = True Then objHTTPemail.setProxyCredentials proxyUsername, proxyPassword
        objHTTPemail.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        objHTTPemail.setTimeouts 100000, 100000, 100000, 100000
        objHTTPemail.setOption 2, 13056
        objHTTPemail.send (requestStr)
        Debug.Print requestStr
        authResponse = objHTTPemail.responsetext
        Debug.Print authResponse
        Set objHTTPemail = Nothing
    End If

    If InStr(1, authResponse, "ERROR->") > 0 Then
        Call checkE(email, dataSource)
    End If


    refreshToken = oldToken


    Exit Function

errhandler:

    stParam2 = "OAUTHREFRESHERROR|" & Err.Number & "|" & Err.Description
    Call checkE(email, dataSource, True)
    Resume Next

End Function



Public Function shortenURL(ByVal urlToShorten As String) As String
    On Error Resume Next
    Dim URL As String
    Dim requestStr As String
    Dim resultURL As String

    URL = "https://supermetrics.com/api/shortenURL?url=" & uriEncode(urlToShorten) & "&responseFormat=RSCL"


    resultURL = vbNullString
    If usingMacOSX = True Or useQTforDataFetch = True Then
        Call fetchDataWithQueryTableDirect(URL, "")
        resultURL = queryTableResultStr
    Else
        Dim xml As Object
        Call setMSXML(xml)
        xml.Open "GET", URL, False
        xml.setTimeouts 20000, 20000, 20000, 20000
        xml.send (requestStr)
        resultURL = xml.responsetext
    End If
    shortenURL = resultURL


End Function


