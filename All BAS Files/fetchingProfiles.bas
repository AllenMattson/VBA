Attribute VB_Name = "fetchingProfiles"
Option Private Module
Option Explicit



Sub OAuthLogin()

    Application.ScreenUpdating = False

    Call setDatasourceVariables

    Call checkOperatingSystem
    If usingMacOSX = False Then
        ProgressBox.Show False
        ProgressBox.stopButton.Visible = False
    End If
    Call updateProgress(4, "Supermetrics Data Grabber is authenticating you to " & serviceName & ". You should see an authentication page opening in your Internet browser in just a moment.")
    Call testConnection(True)

    DoEvents
    Call openOAuthAuthorizationPage

    aika = Timer
    Range("oauthtoken").value = authToken
    Range("authtoken").value = authToken
    If Left(email, 5) = "Error" Then
        stParam1 = "OAERROR"
        stParam2 = email
        Call checkE(email, dataSource, False)
        MsgBox "Unfortunately, the authentication process failed at some point. Please try again. There may be a temporary issue with Google or Supermetrics servers which prevents the tool from working; it may work if you try again later." & vbCrLf & vbCrLf & "The error message is " & email
        End
    End If

    If dataSource = "GA" Then
        Sheets("cred").Cells(1, 1).value = email
        Sheets("cred").Cells(2, 1).value = "oauth"
    ElseIf dataSource = "AW" Then
        Sheets("cred").Cells(3, 1).value = email
        Sheets("cred").Cells(4, 1).value = "oauth"
        Sheets("cred").Cells(12, 1).value = usernameDisp
    ElseIf dataSource = "AC" Then
        Sheets("cred").Cells(8, 1).value = email
        Sheets("cred").Cells(8, 2).value = usernameDisp
        Sheets("cred").Cells(9, 1).value = "oauth"
    ElseIf dataSource = "FB" Then
        Sheets("cred").Cells(10, 1).value = email
        Sheets("cred").Cells(11, 1).value = usernameDisp
    ElseIf dataSource = "YT" Then
        Sheets("cred").Cells(13, 1).value = email
        Sheets("cred").Cells(14, 1).value = usernameDisp
    ElseIf dataSource = "TW" Then
        Sheets("cred").Cells(15, 1).value = email
        Sheets("cred").Cells(16, 1).value = usernameDisp
    ElseIf dataSource = "ST" Then
        Sheets("cred").Cells(17, 1).value = email
        Sheets("cred").Cells(18, 1).value = usernameDisp
    ElseIf dataSource = "FA" Then
        Sheets("cred").Cells(19, 1).value = email
        Sheets("cred").Cells(20, 1).value = usernameDisp
    ElseIf dataSource = "GW" Then
        Sheets("cred").Cells(21, 1).value = email
        Sheets("cred").Cells(22, 1).value = usernameDisp
    ElseIf dataSource = "MC" Then
        Sheets("cred").Cells(23, 1).value = email
        Sheets("cred").Cells(24, 1).value = usernameDisp
    ElseIf dataSource = "MC" Then
        Sheets("cred").Cells(25, 1).value = email
        Sheets("cred").Cells(26, 1).value = usernameDisp
    End If

    Call checkE(email, dataSource)
    If debugMode Then Debug.Print "Demo version=" & demoVersion

    If dataSource <> "TW" Then
        Call fetchProfileList
    Else
        Call storeLoginToSheet
        Call storeTokenToSheet("Twitter", authToken, email)
        Call prepareConfigSheetAfterLogin
    End If
    Application.StatusBar = False

End Sub

Sub prepareConfigSheetAfterLogin()
    On Error Resume Next
    With configsheet

        stParam1 = "3.02"

        Call unprotectSheets

        stParam1 = "3.0201"
        If usernameDisp = vbNullString Then usernameDisp = email
        configsheet.Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Logged in with user: " & usernameDisp
        Modules.Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Logged in with user: " & usernameDisp
        stParam1 = "3.0202"
        Range("loggedin" & varsuffix).value = True
        stParam1 = "3.0203"

        With configsheet

            .Visible = xlSheetVisible
            .Select


            stParam1 = "3.021"

            With Modules
                .Shapes("loginButton" & varsuffix).Visible = False
                .Shapes("loginButtonArrow" & varsuffix).Visible = False
                .Shapes("loginBoxNote" & varsuffix).Visible = False
                .Shapes("logoutButton" & varsuffix).Visible = True
                .Shapes("authStatusBox" & varsuffix).Visible = True
                .Shapes("licenseNote" & varsuffix).Visible = True
                .Shapes("buttonFC" & varsuffix).Visible = False
            End With


        End With


        stParam1 = "3.03"


        Call unprotectSheets

        If dataSource <> "TW" Then

            If dataSource = "GA" Then
                Call updateVisibilityOfDropdowns("drdga")
                Call updateVisibilityOfDropdowns("drsdga")
                Call updateVisibilityOfDropdowns("drmga")
            ElseIf dataSource = "AW" Then
                Call updateVisibilityOfDropdowns("drdaw")
                Call updateVisibilityOfDropdowns("drsdaw")
                Call updateVisibilityOfDropdowns("drmaw")
            ElseIf dataSource = "AC" Then
                Call updateVisibilityOfDropdowns("drdac")
                Call updateVisibilityOfDropdowns("drsdac")
                Call updateVisibilityOfDropdowns("drmac")
            ElseIf dataSource = "GW" Then
                Call updateVisibilityOfDropdowns("drdgw")
                Call updateVisibilityOfDropdowns("drsdgw")
                Call updateVisibilityOfDropdowns("drmgw")
            ElseIf dataSource = "FB" Then
                Call updateVisibilityOfDropdowns("drdfb")
                Call updateVisibilityOfDropdowns("drsdfb")
                Call updateVisibilityOfDropdowns("drmfb")
            ElseIf dataSource = "YT" Then
                Call updateVisibilityOfDropdowns("drdyt")
                Call updateVisibilityOfDropdowns("drsdyt")
                Call updateVisibilityOfDropdowns("drmyt")
            ElseIf dataSource = "MC" Then
                Call updateVisibilityOfDropdowns("drdmc")
                Call updateVisibilityOfDropdowns("drsdmc")
                Call updateVisibilityOfDropdowns("drmmc")
            ElseIf dataSource = "ST" Then
                Call updateVisibilityOfDropdowns("drdst")
                Call updateVisibilityOfDropdowns("drsdst")
                Call updateVisibilityOfDropdowns("drmst")
            ElseIf dataSource = "MC" Then
                Call updateVisibilityOfDropdowns("drdmc")
                Call updateVisibilityOfDropdowns("drsdmc")
                Call updateVisibilityOfDropdowns("drmmc")
            ElseIf dataSource = "TA" Then
                Call updateVisibilityOfDropdowns("drdta")
                Call updateVisibilityOfDropdowns("drsdta")
                Call updateVisibilityOfDropdowns("drmta")
            Else
                Call updateVisibilityOfDropdowns("drd" & dataSource)
                Call updateVisibilityOfDropdowns("drsd" & dataSource)
                Call updateVisibilityOfDropdowns("drm" & dataSource)
            End If

            If usingMacOSX Then
                configsheet.Shapes("checkFieldsButton").Visible = True
            Else
                configsheet.Shapes("checkFieldsButton").Visible = False
            End If

            Call dateRangeTypeChange

            stParam1 = "3.04"

        End If

        Call setSingleAccountFormatting

        If demoVersion = True Then
            Call setDemoVersionFormatting
        Else
            Call removeDemoVersionFormatting
        End If


        stParam1 = "3.13"

        Call unprotectSheets
        Call showAutomationButtons
        settingsSh.Visible = xlSheetVisible

        stParam1 = "3.14"


        stParam1 = "3.1401"
        'add purchase link

        .Hyperlinks.Add Anchor:=.Shapes("demoNote" & varsuffix), Address:="https://supermetrics.com/product/supermetrics-data-grabber/?" & dataSource & "acc=" & uriEncode(email) & "&appid=" & appID & "&ds=" & dataSource & "&a=purchase" & "&goto=pricing"


        stParam1 = "3.1402"
        Call determineMainFont



    End With




    stParam1 = "3.142"
    '   If dataSource = "GA" Then Call setSampleQueryToCQ

    stParam1 = "3.15"


    Debug.Print "PROFILE LIST TIME: " & Timer - aika

    stParam1 = "7"
    stParam2 = "PR|" & profileCount
    'Call checkE(email, dataSource, True)


    Application.StatusBar = False
    Call hideProgressBox
    Call protectSheets
    Call eraseObjHTTPs

    Application.EnableEvents = True


    stParam1 = "3.16"

    If dataSource <> "TW" Then

        If dataSource = "GA" Then
            Range("profileselections").Cells(1, 1).value = "X"
            '        If usingMacOSX = True Then
            '            With instructionsBox.macwarning
            '                .Visible = True
            '                .Caption = "MAC USERS: Due to bugs in Mac Excel, using the tool on Mac may sometimes lead to strange errors. Restarting Excel usually helps in these cases. Save often."
            '            End With
            '        End If
            instructionsBoxAW.Show False
            '  instructionsBox.Show False
            DoEvents
        ElseIf dataSource = "AW" Then
            Range("profileselectionsAW").Cells(1, 1).value = "X"
            If creatingClientFiles = False Then instructionsBoxAW.Show False
            DoEvents
        ElseIf dataSource = "AC" Then
            Range("profileselectionsAC").Cells(1, 1).value = "X"
            instructionsBoxAW.Show False
            DoEvents
        End If

        Call updateProfileSelections
        stParam4 = vbNullString
    End If

    Call protectSheets
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub fetchProfileList()

    On Error GoTo errhandler
    If debugMode = True Then On Error GoTo 0
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Call unprotectSheets
    Application.EnableEvents = False
    Call checkOperatingSystem
    Call hideMacroInstructions
    Call getProxySettingsIfNeeded

    clientLoginModeForGA = Range("clientLoginModeForGA").value

    Call setDatasourceVariables

    loginType = "PRIMARY"

    stParam1 = "1"
    stParam2 = CStr(usingMacOSX)

    Dim rivi As Long

    stParam1 = "1.01"

    Call eraseObjHTTPs

    stParam1 = "1.02"
    Dim accountsToFetchArr() As Variant
    ReDim accountsToFetchArr(1 To 9)
    Dim segmentsListStartCol As Long



    With configsheet
        If .FilterMode Then .ShowAllData
    End With


    Set profileListStart = Range("profileListStart" & varsuffix)

    segmentsListStartCol = Range("segmentsListStart").Column


    Dim i As Long

    stParam1 = "1.03"


    Dim profileListStartRow As Long
    Dim profileListStartColumn As Long

    profileListStartRow = profileListStart.row
    profileListStartColumn = profileListStart.Column


    Dim vrivi As Long

    stParam1 = "1.04"

    'clear old profiles
    Application.EnableEvents = False

    Range("profiles" & varsuffix).ClearContents
    stParam1 = "1.041"
    Range("profiles" & varsuffix).Interior.ColorIndex = configSheetBackgroundColorIndex

    stParam1 = "1.042"
    If dataSource = "GA" Then Sheets("vars").Range("segments2").ClearContents

    stParam1 = "1.05"


    If usernameDisp = "" Then usernameDisp = email


    stParam1 = "1.06"
    stParam2 = "Fetching " & dataSource & " profiles for " & email


    stParam1 = "1.51"

    If usingMacOSX = False Then ProgressBox.Show False

    Call updateProgress(3, "Authenticating to " & serviceName & "...")

    profileCount = 0

    stParam1 = "1.53"
    stParam4 = authToken

    If Left$(authToken, Len("Error: ")) = "Error: " Then
        If InStr(1, authToken, "BadAuthentication") > 0 Then
            Application.StatusBar = False
            If InStr(1, authToken, "Info=InvalidSecondFactor") > 0 Then
                MsgBox "Logging into " & serviceName & " failed, because Google's 2 step authentication process has been enabled for your account. Supermetrics Data Grabber cannot work directly with this feature," _
                     & " but in your Google account's settings, you can generate and application-specific password that can be used in Supermetrics Data Grabber. See your Google Account's control panel for more information."
                configsheet.Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Application-specific password needed"
                stParam1 = "2.01"
                stParam2 = "2-STEPAUTHREQ"
            Else
                MsgBox "The login credentials are incorrect, cannot authenticate to " & serviceName & "."
                configsheet.Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Incorrect email or password"
                stParam1 = "2"
                stParam2 = "BADAUTH"
            End If
            Call hideProgressBox
            Call protectSheets


            Call checkE(email, dataSource, True)
            Call logout
            Modules.Select
            configsheet.Visible = xlSheetVeryHidden
            End

        Else
            Application.StatusBar = False
            configsheet.Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Authentication failed"
            MsgBox "Unfortunately, authenticating to " & serviceName & " with account " & email & " failed. This may be due to:" & vbCrLf & vbCrLf _
                 & "  -No Internet connection being available" & vbCrLf & vbCrLf _
                 & "  -Excel not being able to connect to the Internet due to a proxy server or a firewall blocking the connection." & vbCrLf & vbCrLf _
                 & "  -A temporary problem in " & serviceName & " - the connection may work if you try again later." & vbCrLf & vbCrLf _
                 & "  -Incorrectly entered login credentials" & vbCrLf & vbCrLf _
                 & "The error message is: " & vbCrLf & vbCrLf & authToken
            Call hideProgressBox
            Call protectSheets
            stParam1 = "2"
            stParam2 = "AUTHERROR: " & authToken
            Call checkE(email, dataSource, True)
            End
        End If
    End If

    stParam1 = "2.01"

    Select Case authToken
    Case "Error: Authentication failed"

        configsheet.Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Authentication failed"
        MsgBox "Unfortunately, authenticating to " & serviceName & " failed. This may be due to incorrectly entered login credentials, but can also be caused by Excel not being able to connect to the Internet (for example, due to a proxy server or a firewall blocking the connection)." & vbCrLf & vbCrLf & "If your internet connection is routed though a proxy server, you can specify the server address and login credentials on the Analytics sheet." & vbCrLf & vbCrLf & "This may also be due to a temporary problem in " & serviceName & " - please try again later."
        Call hideProgressBox
        Call logout
        Modules.Select
        configsheet.Visible = xlSheetVeryHidden
        stParam1 = "3"
        stParam2 = "AUTHFAILED"
        Call checkE(email, dataSource, True)
        Call protectSheets
    Case Else

        Range("authToken" & varsuffix).value = authToken
        stParam1 = "3.01"

        'FETCH GA SEGMENTS
        If dataSource = "GA" Then

            stParam1 = "3.011"
            Call updateProgress(20, "Fetching list of advanced segments...")

            stParam1 = "3.012"

            Range("segments2").ClearContents

            stParam1 = "3.013"
            arr = getAccountData(authToken, "segments", True)

            If IsArray(arr) = False Then
                authToken = refreshToken(authToken)
                Range("authToken").value = authToken
                arr = getAccountData(authToken, "segments")
            End If

            If IsArray(arr) = False Then
                stParam1 = "3.0131"
                Debug.Print "FETCHING SEGMENT LIST FAILED, EMPTY RESULT ARRAY"
                stParam2 = "SEGMERROR|" & arr(1, 1)
                Call checkE(email, dataSource, True)
            Else

                If Left$(arr(1, 1), 6) = "Error:" Then
                    authToken = refreshToken(authToken)
                    Range("authToken").value = authToken
                    arr = getAccountData(authToken, "segments")
                End If
                If Left$(arr(1, 1), 6) = "Error:" Then arr = getAccountData(authToken, "segments")

                If Left$(arr(1, 1), 6) = "Error:" Then
                    stParam4 = arr(1, 1)
                    stParam1 = "3.014"
                    Debug.Print "FETCHING SEGMENT LIST FAILED: " & arr(1, 1)
                    stParam1 = "4"
                    stParam2 = "SEGMERROR|" & arr(1, 1)
                    Call checkE(email, dataSource, True)
                Else
                    With Sheets("vars")
                        stParam4 = arr(1, 1)
                        stParam1 = "3.015"
                        For rivi = 1 To UBound(arr)
                            .Cells(rivi + 7, segmentsListStartCol).value = arr(rivi, 1) & "   (id: " & arr(rivi, 2) & ")"
                            .Cells(rivi + 7, segmentsListStartCol + 2).value = arr(rivi, 2)
                        Next rivi
                        stParam1 = "3.016"


                        With .Range(.Cells(8, segmentsListStartCol), .Cells(vikarivi(.Cells(1, segmentsListStartCol)), segmentsListStartCol))
                            .Name = "segments"
                            With .Resize(, 3)
                                .Name = "segments2"
                                .Offset(12).sort key1:=.Cells(1, 1), order1:=xlAscending
                            End With
                        End With

                    End With
                End If

            End If

            'GOALS
            arr = getAccountData(authToken, "goals", True)
            If IsArray(arr) = True Then
                If Left(arr(1, 1), 6) <> "Error:" Then
                    With Range("goalsListStart").Resize(UBound(arr), 3)
                        .value = arr
                        .Name = "goals"
                    End With
                End If
            End If

        End If

        stParam1 = "3.019"

        Call getAccountDataOuter(profileListStart, "PRIMARY")
        Call storeLoginToSheet

        Call prepareConfigSheetAfterLogin

    End Select


    Exit Sub


errhandler:

    stParam2 = "PROFGENERROR|"
    stParam2 = "PROFGENERROR|" & Err.Number & "|" & Err.Description
    Debug.Print "PROFGENERROR: " & stParam1 & " " & stParam2
    Call checkE(email, dataSource, , True)
    Resume Next

End Sub



Sub getAccountDataOuter(profileListStart As Range, Optional loginType As String = "PRIMARY", Optional refreshing As Boolean = False)

    On Error GoTo errhandler

    Dim arr As Variant

    Application.EnableEvents = False

    Call unprotectSheets

    Set profileListStart = Range("profileListStart" & varsuffix)
    Dim profileListStartRow As Long
    Dim profileListStartColumn As Long

    Dim oldProfilesCount As Long
    Dim accountCount As Long
    Dim lastUsedProfileRow As Long
    Dim rivi As Long

    Dim profCombinedStr As String
    Dim profRowRng As Range

    Dim profName As String
    Dim accountName As String

    Dim errorFoundInProfileArr As Boolean

    If Range("loggedin" & varsuffix).value = True Then
        If refreshing And loginType = "PRIMARY" Then
            oldProfilesCount = 0
        Else
            oldProfilesCount = Range("profiles" & varsuffix).Rows.Count
        End If
    Else
        oldProfilesCount = 0
    End If

    Dim vrivi As Long

    profileListStartRow = profileListStart.row
    profileListStartColumn = profileListStart.Column


    aika = Timer
    progresspct = 25

    If Not refreshing Then
        If dataSource = "FB" Then
            Call updateProgress(progresspct, "Fetching list of Facebook pages and applications...")
        ElseIf dataSource = "YT" Then
            Call updateProgress(progresspct, "Fetching list of channels and videos...")
        Else
            Call updateProgress(progresspct, "Fetching list of accounts... ")
        End If
    End If

    stParam1 = "3.05"



    If refreshing Then
        arr = getAccountData(authToken, "profiles", False, False)
    Else
        arr = getAccountData(authToken, "profiles", True)
    End If



    If Not IsArray(arr) Then

        If dataSource = "AC" And InStr(1, arr, "Authentication failed") > 0 Then
            MsgBox "Logging into Bing Ads failed for user " & email & ". The username or password is incorrect. The error message is:" & vbCrLf & vbCrLf & arr
            Call hideProgressBox
            If loginType = "PRIMARY" Then
                Call logout
                Modules.Select
                configsheet.Visible = xlSheetVeryHidden
            End If
            Call protectSheets
            End
        End If

        If Not refreshing Then Call updateProgress(26, "Fetching list of accounts... ")
        authToken = refreshToken(authToken)
        If loginType = "PRIMARY" Then Range("authToken").value = authToken
        If refreshing Then
            arr = getAccountData(authToken, "profiles", False, False)
        Else
            arr = getAccountData(authToken, "profiles", True)
        End If
    End If
    If Not IsArray(arr) Then
        stParam1 = "3.0531"
        If Not refreshing Then Call updateProgress(27, "Fetching list of accounts... ")
        Application.Wait Now + TimeValue("00:00:03")
        arr = getAccountData(authToken)
        If Not IsArray(arr) Then
            stParam1 = "3.0532"
            If loginType = "PRIMARY" Then configsheet.Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Fetching account data failed"
            MsgBox "Unfortunately, fetching your account data failed. This may be due to a temporary problem in " & serviceName & ". The error message is:" & vbCrLf & vbCrLf & arr(1, 1)
            Call hideProgressBox
            If loginType = "PRIMARY" Then
                Call logout
                Modules.Select
                configsheet.Visible = xlSheetVeryHidden
            End If
            Call protectSheets
            stParam2 = "PROFERROR" & loginType & "|" & arr(1, 1)
            Call checkE(email, dataSource, True)
            End
        End If
    End If
    stParam1 = "3.0534"
    If Left$(arr(1, 1), 6) = "Error:" Then
        authToken = refreshToken(authToken)
        If loginType = "PRIMARY" Then Range("authToken").value = authToken
        If Not refreshing Then Call updateProgress(28, "Fetching list of accounts... ")
        If refreshing Then
            arr = getAccountData(authToken, "profiles", False, False)
        Else
            arr = getAccountData(authToken, "profiles", True)
        End If
    End If
    If Left$(arr(1, 1), 6) = "Error:" Then
        If Not refreshing Then Call updateProgress(29, "Fetching list of accounts... ")
        If refreshing Then
            arr = getAccountData(authToken, "profiles", False, False)
        Else
            arr = getAccountData(authToken, "profiles", True)
        End If
    End If



    stParam1 = "3.06"
    errorFoundInProfileArr = False
    If Not IsArray(arr) Then
        errorFoundInProfileArr = True
    ElseIf Left$(arr(1, 1), 6) = "Error:" Then
        errorFoundInProfileArr = True
    End If

    If usernameDisp = "" Then usernameDisp = email

    If errorFoundInProfileArr = True Then
        stParam4 = arr(1, 1)
        If InStr(1, arr(1, 1), "User does not have permissions to use AdWords API") > 0 Then
            stParam1 = "3.061"
            If loginType = "PRIMARY" Then configsheet.Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "User's access rights insufficient"
            MsgBox "Unfortunately, fetching your account data failed. It appears that your user account has insufficient access rights with the AdWords MCC account; Standard or Administrative access level is required." & vbCrLf & vbCrLf & "The error message is:" & vbCrLf & vbCrLf & arr(1, 1)
        ElseIf InStr(1, arr(1, 1), "INCOMPLETE_SIGNUP_LATEST_ADWORDS_API_TNC_NOT_AGREED") > 0 Then
            stParam1 = "3.062"
            If loginType = "PRIMARY" Then configsheet.Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "AdWords billing status invalid"
            MsgBox "Unfortunately, fetching your account data failed. It appears that there's a billing problem with the MCC account. Before starting to use the AdWords API, you need to approve the Terms & Conditions on the Billing settings page of the MCC account. " & vbCrLf & vbCrLf & "The error message is:" & vbCrLf & vbCrLf & arr(1, 1)
        ElseIf InStr(1, arr(1, 1), "No accounts found") > 0 Or InStr(1, arr(1, 1), "No accounts found") > 0 Or InStr(1, arr(1, 1), "No clients found") > 0 Then
            stParam1 = "3.063"
            If loginType = "PRIMARY" Then configsheet.Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "No accounts found"
            MsgBox "No " & serviceName & " accounts were found for user " & usernameDisp & ". Are you sure this is the account you use for " & serviceName & "?"
        Else
            stParam1 = "3.064"
            If loginType = "PRIMARY" Then configsheet.Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Fetching account data failed"
            MsgBox "Unfortunately, fetching your " & serviceName & " account data for user " & usernameDisp & " failed. This may be due to a temporary problem in " & serviceName & ". The error message is:" & vbCrLf & vbCrLf & arr(1, 1)
        End If
        Call hideProgressBox
        If loginType = "PRIMARY" Then
            Call logout
            Modules.Select
            configsheet.Visible = xlSheetVeryHidden
        End If
        stParam2 = "PROFERROR" & loginType & "|" & arr(1, 1)
        Call checkE(email, dataSource, True)
        Call protectSheets
        End
    End If

    stParam1 = "3.07"




    progresspct = 90
    If Not refreshing Then Call updateProgress(progresspct, "Formatting...")


    stParam1 = "3.08"


    lastUsedProfileRow = profileListStartRow - 1 + oldProfilesCount

    With profileListStart.Worksheet

        .Unprotect




        stParam1 = "3.082"
        profileCount = 0
        For rivi = 1 To UBound(arr)

            If CStr(arr(rivi, 3)) <> vbNullString Then
                stParam1 = "3.0821"

                If rivi + lastUsedProfileRow >= rowLimit And rivi > 1000 Then Exit For
                stParam1 = "3.0822"
                profileCount = profileCount + 1


                Set profRowRng = .Cells(rivi + lastUsedProfileRow, profileListStartColumn - 1).Resize(1, 4)
                If dataSource = "FB" Then
                    .Cells(rivi + lastUsedProfileRow, profileListStartColumn).value = arr(rivi, 1)     'type
                    .Cells(rivi + lastUsedProfileRow, profileListStartColumn + 1).value = arr(rivi, 2)  'prof name
                    .Cells(rivi + lastUsedProfileRow, profileListStartColumn + 2).value = arr(rivi, 3)  'profid
                    .Cells(rivi + lastUsedProfileRow, profileListStartColumn - 1).value = usernameDisp    'login

                    profRowRng.Cells(1, 1).value = usernameDisp    'login
                    profRowRng.Cells(1, 2).value = arr(rivi, 1)    'type
                    profRowRng.Cells(1, 3).value = arr(rivi, 2)  'prof name
                    profRowRng.Cells(1, 4).value = arr(rivi, 3)  'profid

                    profName = arr(rivi, 2)
                    accountName = profName

                    '1 type
                    '2 name
                    '3 ID
                    '4 category
                    profCombinedStr = arr(rivi, 2) & " (" & arr(rivi, 3) & ")"
                Else
                    profRowRng.Cells(1, 2).value = arr(rivi, 1)  'account name
                    If arr(rivi, 2) = vbNullString Then
                        profRowRng.Cells(1, 3).ClearContents
                    Else
                        profRowRng.Cells(1, 3).value = arr(rivi, 2)  'prof name
                    End If
                    profRowRng.Cells(1, 4).value = arr(rivi, 3)  'profid
                    profRowRng.Cells(1, 1).value = usernameDisp    'login
                    profCombinedStr = arr(rivi, 1) & ": " & arr(rivi, 2) & " (" & arr(rivi, 3) & ")"


                    If dataSource = "YT" Then
                        profName = arr(rivi, 1)
                        accountName = parseUserName(usernameDisp)
                    Else
                        profName = arr(rivi, 2)
                        accountName = arr(rivi, 1)
                    End If

                End If


                .Hyperlinks.Add Anchor:=profRowRng, Address:="", ScreenTip:=profCombinedStr
                With profRowRng.Font
                    .ColorIndex = 1
                    .Underline = False
                End With


                ' With profRowRng.Offset(, -1).Resize(1, 5)
                ' .Interior.ColorIndex = 50
                ' .Font.ColorIndex = 2
                ' If loginType = "PRIMARY" Then .Cells(1, 2).Font.ColorIndex = 16
                ' End With
                stParam1 = "3.0823"
                Call storeTokenToSheet(arr(rivi, 3), authToken, email, profName, accountName)

                If rivi Mod 10 = 1 Then Call updateProgressIterationBoxes

            End If

        Next rivi

        Call updateProgressIterationBoxes("EXITLOOP")

        stParam1 = "3.09"

        Call eraseObjHTTPs

        Call unprotectSheets

        lastUsedProfileRow = vikarivi(.Cells(1, profileListStartColumn + 1))
        Application.ScreenUpdating = True
        Application.ScreenUpdating = False

        stParam1 = "3.10"


        If profileCount = 0 Then profileCount = 1

        rivi = vikarivi(profileListStart.Offset(, 2)) - profileListStart.row + 1
        If rivi > profileCount Then profileCount = rivi

        profileListStart.Offset(0, -2).Resize(profileCount, 5).Name = "profiles" & varsuffix
        profileListStart.Offset(0, -2).Resize(profileCount, 1).Name = "profileSelections" & varsuffix




        stParam1 = "3.11"


        If dataSource <> "YT" And dataSource <> "MC" Then Range("profiles" & varsuffix).sort key1:=profileListStart, key2:=profileListStart.Offset(0, 1), order1:=xlAscending, order2:=xlAscending

        progresspct = 94
        If Not refreshing Then Call updateProgress(progresspct, "Formatting...")

        Call unprotectSheets

        Call addProfileSelectionCheckBoxes(oldProfilesCount)

        progresspct = 95
        If Not refreshing Then Call updateProgress(progresspct, "Formatting...")


        Call clearProfileSelections
        Call updateProfileSelections

        stParam1 = "3.12"

        Range("profiles" & varsuffix).Resize(1000).Locked = True

        With Range("profiles" & varsuffix)
            .Locked = False
            .NumberFormat = "@"
            .Font.Size = 8
            .HorizontalAlignment = xlLeft
        End With

        If dataSource = "YT" Then
            profileListStart.Offset(0, 1).Resize(profileCount, 1).NumberFormatLocal = Range("numformatDate").NumberFormatLocal
        End If

        Range("profileSelections" & varsuffix).HorizontalAlignment = xlCenter

        stParam1 = "3.121"

        rivi = Range("profiles" & varsuffix).row + Range("profiles" & varsuffix).Rows.Count
        stParam1 = "3.1211"
        If rivi < 50 Then rivi = 50
        .ScrollArea = "A1:Z" & rivi + 30
        stParam1 = "3.1212"

        If profileListStart.row + profileCount < 50 Then
            rivi = 55
        Else
            rivi = profileListStart.row + profileCount + 5
        End If
        stParam1 = "3.1213"

    End With

    Exit Sub


errhandler:


    stParam2 = "ADATAGENERROR" & loginType & "|" & Err.Number & "|" & Err.Description
    Debug.Print "ADATAGENERROR: " & stParam1 & " " & stParam2
    Call checkE(email, dataSource, , True)
    Resume Next

End Sub


Sub refreshSegmentList()

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Application.ScreenUpdating = False
    Dim arr As Variant
    Dim rivi As Long
    Dim segmentsListStartCol As Long
    Dim varArr As Variant

    dataSource = "GA"

    Call checkOperatingSystem
    Call getProxySettingsIfNeeded

    Application.StatusBar = "Refreshing segment list... 14 %"

    filterUF.refreshingNote.Visible = True

    segmentsListStartCol = Range("segmentsListStart").Column

    email = getFirstLoginEmail()
    authToken = getTokenForEmail(email)
    arr = getAccountData(authToken, "segments", , False)

    Application.StatusBar = "Refreshing segment list... 78 %"

    If Left$(arr(1, 1), 6) = "Error:" Then
        stParam4 = arr(1, 1)
        stParam1 = "3.014"
        Debug.Print "FETCHING SEGMENT LIST FAILED: " & arr(1, 1)
        stParam1 = "4"
        stParam2 = "SEGMERROR|" & arr(1, 1)
        Call checkE(email, dataSource, True)

    Else
        With Sheets("vars")
            Range("segments2").ClearContents
            stParam4 = arr(1, 1)
            stParam1 = "3.015"
            For rivi = 1 To UBound(arr)
                .Cells(rivi + 7, segmentsListStartCol).value = arr(rivi, 1) & "   (id: " & arr(rivi, 2) & ")"
                .Cells(rivi + 7, segmentsListStartCol + 2).value = arr(rivi, 2)
            Next rivi
            stParam1 = "3.016"
            With .Range(.Cells(8, segmentsListStartCol), .Cells(vikarivi(.Cells(1, segmentsListStartCol)), segmentsListStartCol))
                .Name = "segments"
                With .Resize(, 3)
                    .Name = "segments2"
                    '       .Offset(12).sort key1:=.Cells(1, 1), order1:=xlAscending
                End With
            End With
        End With

        Application.StatusBar = "Refreshing segment list... 94 %"

        filterUF.Controls("segmentLB").Clear
        varArr = Range("segments").value
        For rivi = 1 To UBound(varArr)
            filterUF.Controls("segmentLB").AddItem (varArr(rivi, 1))
        Next rivi

    End If
    filterUF.refreshingNote.Visible = False
    Application.StatusBar = "Refreshing segment list done"

    Application.StatusBar = False

End Sub



Sub addLogin()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Application.ScreenUpdating = False

    Dim arr As Variant

    Call setDatasourceVariables

    loginType = "SECONDARY"
    ' Call testConnection

    If usingMacOSX = False Then ProgressBox.Show False

    If dataSource <> "AC" Then

        Call updateProgress(3, "Supermetrics Data Grabber is authenticating you to " & serviceName & ". Authentication page should open in your browser in just a moment.")
        DoEvents
        Call openOAuthAuthorizationPage
        If Left(email, 5) = "Error" Then
            stParam1 = "OAERROR"
            stParam2 = email
            Call checkE(email, dataSource, False)
            MsgBox "Unfortunately, the authentication process with " & serviceName & " failed at some point. Please try again. There may be a temporary issue with " & serviceName & " or Supermetrics servers which prevents the tool from working; it may work if you try again later." & vbCrLf & vbCrLf & "The error message is " & email, , "Authentication failed"
            Call protectSheets
            End
        End If

    Else
        Call updateProgress(3, appName & " is authenticating you to " & serviceName & ".")
        BingLoginTypeChoice.Show
        With loginBox
            .emailInput.Text = vbNullString
            .pwInput.Text = vbNullString
            .Caption = "Adding another login: please log into " & serviceName & " with the account you wish to add"

        End With


        If email = vbNullString Or password = vbNullString Then End

    End If

    If usernameDisp = vbNullString Then usernameDisp = email

    If findRowWithValue(loginInfoCol, "em$" & email, 1, Sheets("logins"), 1) > 0 Then
        Call hideProgressBox
        MsgBox "You tried to add account " & usernameDisp & ", but you have already logged into the " & moduleName & " with this user.", , "User already logged in"
        Call protectSheets
        End
    End If
    Call checkE(email, dataSource, False, True)

    If licenseStatus = "INVALID" Then
        Call hideProgressBox
        With buyLicenseBox
            .note1.Caption = "The Supermetrics Data Grabber " & moduleName & " license of the login you tried to add (" & usernameDisp & ") has expired. To add this login, please visit Supermetrics.com to purchase a license." & vbCrLf & vbCrLf & "When making the purchase on the site, use this user ID:"
            .usernameTB.Text = email
            .Show
        End With
        Call protectSheets
        End
    ElseIf demoStatus = "INVALID" Then
        Call hideProgressBox
        With buyLicenseBox
            .note1.Caption = "The Supermetrics Data Grabber " & moduleName & " trial period of the login you tried to add (" & usernameDisp & ") has expired. To add this login, please visit Supermetrics.com to purchase a license." & vbCrLf & vbCrLf & "When making the purchase on the site, use this user ID:"
            .usernameTB.Text = email
            .Show
        End With
        Call protectSheets
        End
    End If


    Call getProfilesForLogin(authToken)
    Call storeLoginToSheet
    Call setMultiAccountFormatting
    Call updateLoginStatus

    If dataSource = "GA" Then
        'GOALS
        arr = getAccountData(authToken, "goals")
        If IsArray(arr) = True Then
            If Left(arr(1, 1), 6) <> "Error:" Then
                Range("goalsListStart").Offset(Range("goals").Rows.Count).Resize(UBound(arr), 3).value = arr
                Range("goalsListStart").Resize(Range("goals").Rows.Count + UBound(arr), 3).Name = "goals"
            End If
        End If
    End If

    Call hideProgressBox
    configsheet.Select
    MsgBox "Login " & usernameDisp & " added successfully. The profiles and accounts this login has access to are now included in the profile list.", , "Login has been added"
End Sub




Public Function getAccountData(Optional ByVal token As String, Optional dataType As String = "profiles", Optional fromStr As Boolean = False, Optional showProgressBox As Boolean = True) As Variant

' On Error GoTo errhandler
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    stParam1 = "13.051"

    Call checkOperatingSystem

    Dim requestStr As String
    Dim responseStr As String

    'Debug.Print

    Dim URL As String

    Dim rivi As Long

    Dim resultArr As Variant
    Dim tempArr As Variant
    Dim accountNamesArr As Variant
    Dim profileNamesArr As Variant
    Dim profileIDsArr As Variant
    Dim profileCount As Long
    Dim prevResponse As String

    Dim resultsArrCol1 As Variant
    Dim resultsArrCol2 As Variant
    Dim resultsArrCol3 As Variant
    Dim rowCount As Long

    Const maxAccountsPerIteration As Long = 1000

    Dim errorStr As String
    Dim errorCount As Long

    Dim i As Long

    Dim tempStr As String
    Dim connectionErrorCount As Long

    If dataSource = "AC" Then showProgressBox = False

    If fromStr = True Then
        If dataType = "segments" Then
            responseStr = segmentsStr
        ElseIf dataType = "goals" Then
            responseStr = goalsStr
        Else
            responseStr = profilesStr
        End If
    End If

    If fromStr = False Or responseStr = vbNullString Then

        connectionErrorCount = 0

        stParam1 = "13.052"

        responseStr = ""

        For iterationNum = 1 To 1000

            URL = "https://supermetrics.com/api/getAccount?responseFormat=RSCL"

            requestStr = "token=" & uriEncode(token)

            If dataType = "segments" Then
                requestStr = requestStr & "&datatype=segments"
            ElseIf dataType = "goals" Then
                requestStr = requestStr & "&datatype=goals"
            ElseIf dataType = "webproperties" Then
                requestStr = requestStr & "&datatype=webproperties"
            Else
                requestStr = requestStr & "&datatype=profiles"
            End If

            requestStr = requestStr & "&version=" & uriEncode(versionNumber)
            requestStr = requestStr & "&rid=" & randID
            requestStr = requestStr & "&encoding=light"
            requestStr = requestStr & "&arrayType=combined2"
            requestStr = requestStr & "&dataSource=" & dataSource

            requestStr = requestStr & "&startFromAccount=" & (iterationNum - 1) * maxAccountsPerIteration
            requestStr = requestStr & "&endOnAccount=" & iterationNum * maxAccountsPerIteration - 1

            If separatorList = vbNullString Then
                separatorList = "&rscL1=" & uriEncode(rscL1)
                separatorList = separatorList & "&rscL2=" & uriEncode(rscL2)
                separatorList = separatorList & "&rscL3=" & uriEncode(rscL3)
                separatorList = separatorList & "&rscL4=" & uriEncode(rscL4)
                separatorList = separatorList & "&rscL0=" & uriEncode(rscL0)
            End If

            requestStr = requestStr & separatorList

            Debug.Print "Req: " & requestStr


            If debugMode = True Then Debug.Print requestStr
runFetchAgain:
            If usingMacOSX = False And useQTforDataFetch = False Then
                stParam1 = "13.053"
                requestStr = requestStr & "&chrencode=0"
                Call setMSXML(objhttp)
                If useProxy = True Then objhttp.setProxy 2, proxyAddress
                objhttp.Open "POST", URL, True
                If useProxyWithCredentials = True Then objhttp.setProxyCredentials proxyUsername, proxyPassword
                objhttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
                objhttp.setTimeouts 1000000, 1000000, 1000000, 1000000
                On Error GoTo errHandler2
                objhttp.send (requestStr)
                On Error Resume Next
                If debugMode = True Then On Error GoTo 0

                Do
                    objhttp.waitForResponse 0
                    If objhttp.readyState = 4 Then Exit Do
                    If showProgressBox Then Call updateProgressIterationBoxes
                Loop

                tempStr = objhttp.responsetext
                Debug.Print "Resp: " & Left(tempStr, 2000)

            Else
                stParam1 = "13.054"
                requestStr = requestStr & "&chrencode=1"
                requestStr = requestStr & "&urlencode=1"
                Call fetchDataWithQueryTableDirect(URL, requestStr, True, True)
                tempStr = queryTableResultStr
                tempStr = chrDecode(tempStr)
            End If

            If tempStr = prevResponse Then Exit For
            prevResponse = tempStr

            If tempStr = "ALL_ACCOUNTS_FETCHED" Then Exit For
            If iterationNum > 1 Then
                responseStr = responseStr & Right$(tempStr, Len(tempStr) - Len("SUCCESS" & rscL2))
            Else
                responseStr = tempStr
            End If
            If showProgressBox Then
                progresspct = Evaluate("25+65*" & iterationNum & "/(" & iterationNum & " + 20)")
                Call updateProgress(progresspct, "", "Fetched: " & Replace(Left$(Right$(tempStr, Len(tempStr) - 10), 100), rscL4, " ") & "...")
            End If
            If usingMacOSX = False And useQTforDataFetch = False Then Set objhttp = Nothing

            If dataType <> "profiles" Or dataSource <> "GA" Then Exit For
        Next iterationNum

    End If


    stParam1 = "13.055"


    If Left(responseStr, 6) = "Error:" Then
        ReDim resultArr(1 To 1, 1 To 1)
        resultArr(1, 1) = responseStr
        stParam1 = "13.0575"
        getAccountData = resultArr
        Exit Function
    End If




    stParam1 = "13.056"

    If debugMode = True Then Debug.Print Left(responseStr, 5000)

    stParam1 = "13.057"


    If InStr(1, responseStr, "Error: license expired") > 0 Then
        ReDim resultArr(1 To 1, 1 To 1)
        resultArr(1, 1) = "Your license has expired. Visit Supermetrics.com to purchase the full version."
        stParam1 = "13.0575"
        getAccountData = resultArr
        Exit Function
    End If



    stParam1 = "13.058"




    stParam1 = "13.0581"

    If dataSource <> "GA" Then


        resultArr = Split(responseStr, rscL2)

        resultsArrCol1 = Split(resultArr(1), rscL3)  'acc name
        resultsArrCol2 = Split(resultArr(2), rscL3)  'name
        resultsArrCol3 = Split(resultArr(3), rscL3)  'ID

        rowCount = UBound(resultsArrCol3) + 1

        If rowCount = 0 Then
            Debug.Print "No data found"
            ReDim resultArr(1 To 1, 1 To 1)
            resultArr(1, 1) = "Error: No data found"
            getAccountData = resultArr
            Exit Function
        End If

        ReDim resultArr(1 To rowCount, 1 To 3)

        For rivi = 1 To rowCount
        If UBound(resultsArrCol1) >= rivi - 1 Then
            resultArr(rivi, 1) = Left(resultsArrCol1(rivi - 1), 255)
            Else
            resultArr(rivi, 1) = ""
            End If
            If dataSource = "YT" Then
                resultArr(rivi, 3) = Left(resultsArrCol2(rivi - 1), 255)
                resultArr(rivi, 2) = Left(resultsArrCol3(rivi - 1), 255)
            Else
                resultArr(rivi, 2) = Left(resultsArrCol2(rivi - 1), 255)
                resultArr(rivi, 3) = Left(resultsArrCol3(rivi - 1), 255)
            End If
        Next rivi


    ElseIf dataType = "profiles" Or dataType = "webproperties" Then
        stParam1 = "13.058101"
        tempArr = Split(responseStr, rscL3)
        profileCount = UBound(tempArr)
        If profileCount = 0 Then
            stParam1 = "13.058102"
            Debug.Print "No profiles found"
            ReDim resultArr(1 To 1, 1 To 1)
            resultArr(1, 1) = "Error: No clients found"
            getAccountData = resultArr
            Exit Function
        End If
        If dataType = "webproperties" Then
            ReDim resultArr(1 To profileCount, 1 To 4)
        Else
            ReDim resultArr(1 To profileCount, 1 To 3)
        End If
        '1 acc name
        '2 name
        '3 ID
        '4 (account ID)
        For rivi = 1 To profileCount
            resultArr(rivi, 1) = Split(tempArr(rivi), rscL4)(0)
            resultArr(rivi, 2) = Split(tempArr(rivi), rscL4)(1)
            resultArr(rivi, 3) = Split(tempArr(rivi), rscL4)(2)
            If dataType = "webproperties" Then resultArr(rivi, 4) = Split(tempArr(rivi), rscL4)(3)

            If rivi > 1 Then
                If resultArr(rivi, 1) = vbNullString Then resultArr(rivi, 1) = resultArr(rivi - 1, 1)
                If dataType = "webproperties" Then
                    If resultArr(rivi, 4) = vbNullString Then resultArr(rivi, 4) = resultArr(rivi - 1, 4)
                End If
            End If
        Next rivi
    Else


        stParam1 = "13.05811"
        resultArr = Split(responseStr, rscL2)

        accountNamesArr = Split(resultArr(1), rscL3)   'goals: profid
        profileNamesArr = Split(resultArr(2), rscL3)   'goals: goalid
        If dataType <> "segments" Then profileIDsArr = Split(resultArr(3), rscL3)   'goals: goalname

        stParam1 = "13.059"
        profileCount = UBound(profileNamesArr) + 1

        stParam1 = "13.0593"

        If profileCount = 0 Then
            stParam1 = "13.0594"
            Debug.Print "No profiles found"
            ReDim resultArr(1 To 1, 1 To 1)
            resultArr(1, 1) = "Error: No clients found"
            getAccountData = resultArr
            Exit Function
        End If


        ReDim resultArr(1 To profileCount, 1 To 3)
        '1 acc name
        '2 name
        '3 ID

        For rivi = 1 To profileCount
            resultArr(rivi, 1) = accountNamesArr(rivi - 1)
            resultArr(rivi, 2) = profileNamesArr(rivi - 1)
            If dataType <> "segments" Then
                resultArr(rivi, 3) = profileIDsArr(rivi - 1)
                If rivi > 1 Then
                    If resultArr(rivi, 1) = vbNullString Then resultArr(rivi, 1) = resultArr(rivi - 1, 1)
                End If
            End If
        Next rivi
    End If


    resultArr = arrayReplace(resultArr, "%rscL1%", rscL1)
    resultArr = arrayReplace(resultArr, "%rscL2%", rscL2)
    resultArr = arrayReplace(resultArr, "%rscL3%", rscL3)
    resultArr = arrayReplace(resultArr, "%rscL4%", rscL4)

    getAccountData = resultArr


    Exit Function

errhandler:
    Resume Next
    '    ReDim resultArr(1 To 1, 1 To 1)
    '    resultArr(1, 1) = "Error: " & Err.Description
    '    getAccountData = resultArr

    Exit Function

errHandler2:
    Resume runFetchAgain

    Exit Function

End Function



