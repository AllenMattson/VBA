Attribute VB_Name = "proxyConf"
Option Private Module
Option Explicit

Sub getProxySettingsIfNeeded()
    If usingMacOSX = False Then
        proxyAddress = Range("proxyAddress").value
        proxyUsername = Range("proxyUsername").value
        proxyPassword = Range("proxyPassword").value
    End If

    If proxyAddress <> vbNullString Then
        useProxy = True
        If proxyUsername <> vbNullString Then
            useProxyWithCredentials = True
        Else
            useProxyWithCredentials = False
        End If
    Else
        useProxy = False
    End If
End Sub

Sub askForProxyCredentials()

    With proxyBox
        .proxyaddressInput.Text = Range("proxyaddress").value
        .proxyusernameInput.Text = Range("proxyusername").value
        .proxypasswordInput.Text = Range("proxypassword").value
        .Show
    End With

End Sub

Sub testConnection(Optional askForProxyCredentialsBeforeTryingQT As Boolean = False)

    On Error Resume Next

    Application.StatusBar = "Testing internet connection..."

    Dim i As Long

    Call checkOperatingSystem

    If usingMacOSX = True Then Exit Sub

    'test without proxy
    For i = 1 To 2
        If testConnectionWithCurrentSettings(, True) = True Then
            If usingMacOSX = False Then Debug.Print "Connecting with MSXML successful without proxy, using that in data fetch"
            useQTforDataFetch = False
            Range("proxyaddress").value = vbNullString
            Range("useQTforDataFetch").value = False
            Application.StatusBar = False
            Exit Sub
        End If
    Next i
    Debug.Print "Connecting without proxy failed"

    'test with proxy
    Call getProxySettings(False)
    If testConnectionWithCurrentSettings() = True Then GoTo showProxySettings
    Debug.Print "Connecting with proxy but without login credentials failed"

    'test with proxy and stored proxy credentials
    Call getProxySettings(True)
    If testConnectionWithCurrentSettings() = True Then GoTo showProxySettings
    Debug.Print "Connecting with proxy with stored login credentials failed"

    If askForProxyCredentialsBeforeTryingQT Then
        'test with proxy and ask for proxy credentials
        Call askForProxyCredentials
        If testConnectionWithCurrentSettings() = True Then GoTo showProxySettings
        Debug.Print "Connecting with proxy with inputted login credentials failed"
    End If

    If testConnectionWithCurrentSettings(True) = True Then
        Debug.Print "Connecting with QT successful, using that in data fetch"
        Range("useQTforDataFetch").value = True
        useQTforDataFetch = True
        Application.StatusBar = False
        Exit Sub
    End If

    If Not askForProxyCredentialsBeforeTryingQT Then
        'test with proxy and ask for proxy credentials
        Call askForProxyCredentials
        If testConnectionWithCurrentSettings() = True Then GoTo showProxySettings
        Debug.Print "Connecting with proxy with inputted login credentials failed"
    End If

    If testConnectionWithCurrentSettings() = False Then
        MsgBox "Failed to connect to the Internet. This may be due to a firewall blocking the connection. Check your network connection and try again.", , "Network connection error"
        Call hideProgressBox
        Application.StatusBar = False
        End
    End If

    Application.StatusBar = False
    Exit Sub

showProxySettings:
    Sheets("proxysettings").Visible = xlSheetVisible
    Application.StatusBar = False
    Debug.Print "Connecting with MSXML with proxy successful, using that in data fetch"
    Exit Sub

End Sub


Sub getProxySettings(Optional includeCredentials As Boolean = False)

    Dim proxyEnableSetting As String
    Dim proxyServerSetting As String

    proxyEnableSetting = RegKeyRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable")
    proxyServerSetting = RegKeyRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer")

    Debug.Print "ProxyEnabled setting value: " & proxyEnableSetting
    Debug.Print "ProxyServer setting value: " & proxyServerSetting

    If proxyEnableSetting = 1 Then
        Debug.Print "Proxy is enabled in registry"
        proxyAddress = proxyServerSetting

        Range("proxyaddress").value = proxyAddress

        If proxyAddress <> vbNullString Then
            Debug.Print "Proxy address found -> using proxy"
            useProxy = True
            If Range("proxyusername").value <> "" Then
                useProxyWithCredentials = True
                proxyUsername = Range("proxyusername").value
                proxyPassword = Range("proxypassword").value
            End If
        Else
            Debug.Print "Proxy address not found -> not using proxy"
            useProxy = False
        End If
    End If


End Sub


Public Function testConnectionWithCurrentSettings(Optional forceQT As Boolean = False, Optional ignoreProxy As Boolean = False) As Boolean

    testConnectionWithCurrentSettings = False

    On Error Resume Next
    ' If debugMode = True Then On Error GoTo 0

    Dim objHTTPtestConn As Object
    Dim errorCount As Integer
    errorCount = 0

runFetchAgain:
    If usingMacOSX = True Or forceQT = True Then
        On Error GoTo connectionError
        Call fetchDataWithQueryTableDirect("http://www.google.com", "")
        On Error Resume Next
        If debugMode = True Then On Error GoTo 0
        If queryTableResultStr = "" Then
            testConnectionWithCurrentSettings = False
        Else
            testConnectionWithCurrentSettings = True
            If debugMode = True Then Debug.Print "Connection successful: " & queryTableResultStr
        End If
    Else
        Call setMSXML(objHTTPtestConn)
        If ignoreProxy = False And useProxy = True Then objHTTPtestConn.setProxy 2, proxyAddress
        objHTTPtestConn.Open "GET", "http://www.google.com", False
        If ignoreProxy = False And useProxyWithCredentials = True Then objHTTPtestConn.setProxyCredentials proxyUsername, proxyPassword
        objHTTPtestConn.setTimeouts 10000, 10000, 10000, 10000
        On Error GoTo connectionError
        objHTTPtestConn.send ("")
        On Error Resume Next
        If debugMode = True Then On Error GoTo 0
        If debugMode = True Then Debug.Print "Connection successful: " & Left(objHTTPtestConn.responsetext, 200)
        testConnectionWithCurrentSettings = True
    End If


    Set objHTTPtestConn = Nothing




    Exit Function

connectionError:
    If errorCount > 1 Then
        Set objHTTPtestConn = Nothing
        testConnectionWithCurrentSettings = False
    Else
        errorCount = errorCount + 1
        Resume runFetchAgain
    End If
End Function



