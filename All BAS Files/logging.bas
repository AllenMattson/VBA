Attribute VB_Name = "logging"
Option Private Module
Option Explicit





Sub runLogins()

    Dim fname As String
    Dim rivi As Long
    Dim loginWB As Workbook
    Set loginWB = Workbooks("AdWords_aktiiviset asiakkaat_281211_2.xlsx")

    With loginWB.Sheets("Customers")
        For rivi = 1 To 1000

            If .Cells(rivi, 11).value <> vbNullString Then
                ThisWorkbook.Activate
                Debug.Print "Start: " & .Cells(rivi, 11).value
                Call autoLogin(.Cells(rivi, 11).value, .Cells(rivi, 12).value, .Cells(rivi, 13).value)
                Debug.Print "Done: " & .Cells(rivi, 1).value
                fname = .Cells(rivi, 1).value
                fname = Replace(fname, "/", " ")
                fname = Replace(fname, "\", " ")
                fname = Replace(fname, "*", " ")
                Application.DisplayAlerts = False
                ThisWorkbook.SaveAs "C:\Users\Asus Mikael\Documents\Inside e\" & fname & ".xls", 56
                Application.DisplayAlerts = True
            End If

        Next rivi
    End With
End Sub



Sub autoLogin(userName As String, pw As String, profID As Variant)

    Dim rivi As Long
    Dim sar As Long

    creatingClientFiles = True

    Call logoutAW
    Sheets("cred").Cells(3, 1).value = userName
    Sheets("cred").Cells(4, 1).value = pw
    Sheets("cred").Cells(5, 1).value = vbNullString


    dataSource = "AW"
    Call testConnection

    Call fetchProfileList

    Modules.Visible = xlSheetVeryHidden
    AdWords.Visible = xlSheetVeryHidden

    Range("profileliststartaw").Resize(Range("profilesaw").Rows.Count, 1).value = "Menestystarinat"

    rivi = Range("profIDrowQS").row

    With Sheets("querystorage")
        For sar = 1 To 256
            If .Cells(5, sar).value <> vbNullString Then
                .Cells(rivi, sar).value = profID
            End If
        Next sar
    End With

    Call refreshDataOnAllSheetsDontOverrideDates

    creatingClientFiles = False

End Sub


Sub showLoginBoxAC()
    dataSource = "AC"
    BingLoginTypeChoice.Show
    '    loginType = "PRIMARY"
    '    Call updateProgress(5, "Testing network connection... A login box should appear in just a moment.")
    '    Call testConnection
    '    With loginBox
    '        .emailInput.Text = Sheets("cred").Cells(8, 1).value
    '        .pwInput.Text = Sheets("cred").Cells(9, 1).value
    '        .Show
    '    End With
End Sub
Sub showOldLoginBoxAC()

End Sub

'Sub showLoginBoxGW()
'    profilesStr = ""
'    dataSource = "GW"
'    With loginBox
'        .Caption = "Log in to Google Webmaster Tools"
'        .loginNote.Caption = "Please log in to Google Webmaster Tools. Your credentials will only be stored in this file on your own computer."
'        .Show
'    End With
'
'End Sub




Sub logout(Optional askToDestroyTokens As Boolean = False)

    On Error Resume Next

    Call checkOperatingSystem

    Call setDatasourceVariables

    On Error Resume Next
    'If debugMode = True Then On Error GoTo 0
    Application.ScreenUpdating = False
    Call unprotectSheets

    If askToDestroyTokens = True And (dataSource = "GA" Or dataSource = "FB" Or dataSource = "GW" Or dataSource = "FA" Or dataSource = "AW" Or dataSource = "YT" Or dataSource = "TW" Or dataSource = "ST" Or dataSource = "MC" Or dataSource = "TA") Then
        questionUFb1Clicked = False
        questionUFb2Clicked = False
        Call hideProgressBox
        Call showQuestionUF("Do you want to log out of the " & moduleName & " in all Supermetrics Data Grabber files where you're logged in, or just this one? If you choose to log out of all files, then all your access information is deleted from our servers, and you need to reauthenticate to use the tool. Note that this only affects the " & moduleName & ".", "Log out of this file only", "Log out of all SMDG files", "Log out of this file or all SMDG files?")
        If questionUFb1Clicked = False And questionUFb2Clicked = False Then End

        If questionUFb2Clicked = True Then Call destroyTokens
    End If


    Call clearLoginCredentials
    Call clearFieldSelections
    Call clearFilters

    If dataSource = "GA" Then Call setSingleAccountFormatting

    Call deleteDataConnections

    With configsheet

        'clear old profiles
        Application.EnableEvents = False


        With Range("profiles" & varsuffix)
            .Hyperlinks.Delete
            .ClearContents
            .Interior.ColorIndex = configSheetBackgroundColorIndex
        End With



        Application.EnableEvents = True

        If dataSource = "GA" Then
            Sheets("vars").Range("segments2").ClearContents
            Sheets("vars").Range("goals").ClearContents
        End If

        .Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Not logged in"

        .Shapes("manageLoginsButton" & varsuffix).Visible = False
        .Shapes("addLoginButton" & varsuffix).Visible = False
        .Shapes("addLoginButtonNote1" & varsuffix).Visible = False
        .Shapes("addLoginButtonNote2" & varsuffix).Visible = False

        .AutoFilter.ShowAllData



        With Modules
            .Shapes("loginButton" & varsuffix).Visible = True
            .Shapes("loginButtonArrow" & varsuffix).Visible = True
            .Shapes("loginBoxNote" & varsuffix).Visible = True
            .Shapes("logoutButton" & varsuffix).Visible = False
            .Shapes("authStatusBox" & varsuffix).Visible = False
            .Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Not logged in"
            .Shapes("licenseNote" & varsuffix).Visible = False
            .Shapes("licenseNote" & varsuffix).TextFrame.Characters.Text = ""
            .Shapes("buttonFC" & varsuffix).Visible = True

            .Shapes("manageLoginsButton" & varsuffix).Visible = False
            .Shapes("addLoginButton" & varsuffix).Visible = False
            .Shapes("addLoginButtonNote1" & varsuffix).Visible = False
            .Shapes("addLoginButtonNote2" & varsuffix).Visible = False
            .Select
        End With



        Call deleteProfileSelectionCBs

        .Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Not logged in"
        .Shapes("licenseNote" & varsuffix).TextFrame.Characters.Text = ""

        .Visible = xlSheetVeryHidden
    End With


    Sheets("tokens").Cells(1, loginInfoCol).Resize(10000, 6).ClearContents
    Sheets("logins").Cells(1, loginInfoCol).Resize(10000, 6).ClearContents

    Range("advancedSettingsInput" & varsuffix).value = ""
    Range("licenseWarningShown" & varsuffix).value = False
    Range("authToken" & varsuffix).value = ""

    Call removeDemoVersionFormatting

    Range("loggedin" & varsuffix).value = False

    Call protectSheets

End Sub

Sub destroyTokens()
    Dim vrivi As Long
    Dim rivi As Long

    With Sheets("logins")
        vrivi = vikarivi(.Cells(1, 1))
        For rivi = 1 To vrivi
            email = trimEM(.Cells(rivi, loginInfoCol).value)
            authToken = getTokenForEmail(email)

            Dim objHTTPemail As Object
            Dim URL As String
            Dim requestStr As String

            Debug.Print "Destroy tokens for " & email & " " & dataSource

            URL = "https://supermetrics.com/api/destroyTokens?responseFormat=RSCL"

            requestStr = "token=" & authToken
            requestStr = requestStr & "&email=" & email
            requestStr = requestStr & "&appid=" & appID & "&version=" & versionNumber & "&rid=" & randID
            requestStr = requestStr & "&system=" & uriEncode(OSandExcelVersion)

            If usingMacOSX = True Or useQTforDataFetch = True Then
                Call fetchDataWithQueryTableDirect(URL, requestStr)
            Else
                Call setMSXML(objHTTPemail)
                If useProxy = True Then objHTTPemail.setProxy 2, proxyAddress
                objHTTPemail.Open "POST", URL, False
                If useProxyWithCredentials = True Then objHTTPemail.setProxyCredentials proxyUsername, proxyPassword
                objHTTPemail.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
                objHTTPemail.setTimeouts 100000, 100000, 100000, 100000
                objHTTPemail.setOption 2, 13056
                objHTTPemail.send (requestStr)
                Set objHTTPemail = Nothing
            End If
        Next rivi
    End With
End Sub





Sub clearLoginCredentials()
    On Error Resume Next
    If dataSource = "AW" Then
        Sheets("cred").Cells(3, 1).ClearContents
        Sheets("cred").Cells(4, 1).ClearContents
        Sheets("cred").Cells(5, 1).ClearContents
        Range("authtokenaw").value = ""
    ElseIf dataSource = "GA" Then
        Sheets("cred").Cells(1, 1).ClearContents
        Sheets("cred").Cells(2, 1).ClearContents
        Range("authtoken").value = ""
        Range("oauthtoken").value = ""
    ElseIf dataSource = "AC" Then
        Sheets("cred").Cells(8, 1).ClearContents
        Sheets("cred").Cells(9, 1).ClearContents
        Range("authtokenac").value = ""
        loginBox.emailInput.Text = ""
        loginBox.pwInput.Text = ""
    ElseIf dataSource = "FB" Then
        Sheets("cred").Cells(10, 1).ClearContents
        Sheets("cred").Cells(11, 1).ClearContents
        Range("authtokenFB").value = ""
        Range("oauthtokenFB").value = ""
    ElseIf dataSource = "YT" Then
        Sheets("cred").Cells(13, 1).ClearContents
        Sheets("cred").Cells(14, 1).ClearContents
        Range("authtokenYT").value = ""
        Range("oauthtokenYT").value = ""
    ElseIf dataSource = "GW" Then
        Sheets("cred").Cells(16, 1).ClearContents
        Sheets("cred").Cells(17, 1).ClearContents
        Range("authtokenGW").value = ""
        Range("oauthtokenGW").value = ""
    End If
End Sub



Public Function getCLtoken(ByVal email As String, ByVal password As String)

'
'Fetches GA authentication token, which can then be used to fetch data with the getGAdata function
'
'Created by Mikael Thuneberg

    Call checkOperatingSystem
    Dim service As String
    If dataSource = "GW" Then service = "sitemaps"


    Dim CurChr As Long
    Dim tempAns As String
    Dim authRequestStr As String
    Dim authResponse As String
    Dim authTokenStart As Long

    Dim objHTTPauth As Object

    Dim URL As String


    If email = vbNullString Then
        getCLtoken = vbNullString
        Exit Function
    End If

    If password = vbNullString Then password = getPWforEmail(email)

    If password = vbNullString Then
        getCLtoken = "Error: Input password"
        Exit Function
    End If


    URL = "https://www.google.com/accounts/ClientLogin"

    On Error GoTo errhandler
    If debugMode = True Then On Error GoTo 0

    'accountType':'HOSTED_OR_GOOGLE',
    'Email' : email,
    'Passwd': pw,
    'service' : "sitemaps",
    'source' : "Supermetrics"

    authRequestStr = "accountType=HOSTED_OR_GOOGLE&Email=" & uriEncode(email) & "&Passwd=" & uriEncode(password) & "&service=" & service & "&Source=Supermetrics-" & versionNumber


    If usingMacOSX = True Or useQTforDataFetch = True Then
        Call fetchDataWithQueryTableDirect(URL, authRequestStr)
        authResponse = queryTableResultStr
    Else

        Call setMSXML(objHTTPauth)
        If useProxy = True Then objHTTPauth.setProxy 2, proxyAddress
        objHTTPauth.Open "POST", URL, False
        If useProxyWithCredentials = True Then objHTTPauth.setProxyCredentials proxyUsername, proxyPassword
        objHTTPauth.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        objHTTPauth.setTimeouts 100000, 100000, 100000, 100000
        objHTTPauth.setOption 2, 13056
        objHTTPauth.send (authRequestStr)

        authResponse = objHTTPauth.responsetext
        Set objHTTPauth = Nothing
    End If


    If debugMode = True Then Debug.Print "Auth response: " & authResponse

    If InStr(1, authResponse, "InvalidSecondFactor") > 0 Then
        Call showQuestionUF("Authentication error: Your account has 2-step verification enabled, so you need to use an application-specific password. You can create one by clicking the button below.", "Create password", "Cancel", "Application-specific password must be used", False)
        If questionUFb1Clicked Then
            ActiveWorkbook.FollowHyperlink Address:="https://accounts.google.com/b/0/IssuedAuthSubTokens?hide_authsub=1", NewWindow:=True
        End If
        End
    End If

    If InStr(1, authResponse, "Error=CaptchaRequired") > 0 Then
        MsgBox "Error: Captcha required: " & Right$(authResponse, Len(authResponse) - InStr(1, authResponse, "Url=http") - 3)
        End
    End If

    If InStr(1, authResponse, "BadAuthentication") = 0 Then

        authTokenStart = InStr(1, authResponse, "Auth=") + 4
        authToken = Right$(authResponse, Len(authResponse) - authTokenStart)

        authToken = Trim(authToken)
        authToken = Replace(authToken, vbCrLf, "")
        authToken = Replace(authToken, vbCr, "")
        authToken = Replace(authToken, vbLf, "")
        authToken = Trim(authToken)
        getCLtoken = authToken

        '   Call storeToken(authToken, email, appID)
    Else
        MsgBox "Error: Authentication failed " & authResponse
        End

    End If

    Exit Function

errhandler:

    getCLtoken = "Error: Authentication failed " & Err.Description

End Function
