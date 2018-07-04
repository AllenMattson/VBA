Attribute VB_Name = "fetchingProfiles2"
Option Private Module
Option Explicit



Sub refreshProfileList()

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim vrivi As Long
    Dim rivi As Long
    Dim primaryLogin As Boolean
    Dim pctDone As Long

    primaryLogin = True

    Call setDatasourceVariables
    Call checkOperatingSystem
    Call getProxySettingsIfNeeded

    If usingMacOSX = False Then ProgressBox.Show False
    Call updateProgress(4, "Clearing current list...")

    configsheet.AutoFilter.ShowAllData

    Range("profiles" & varsuffix).ClearContents
    Call deleteProfileSelectionCBs
    Range("profiles" & varsuffix).Interior.ColorIndex = 2

    With Sheets("logins")
        vrivi = vikarivi(.Cells(1, loginInfoCol))
        For rivi = 1 To vrivi
            email = trimEM(.Cells(rivi, loginInfoCol).value)
            usernameDisp = .Cells(rivi, loginInfoCol + 4).value
            If email <> vbNullString Then
                If usernameDisp = vbNullString Then usernameDisp = email
                pctDone = 15 + 84 * (rivi - 1) / vrivi
                Call updateProgress(pctDone, "Fetching account list for " & usernameDisp & "...")
                authToken = getTokenForEmail(email)
                Set profileListStart = configsheet.Cells(Range("profileListStart" & varsuffix).row + Range("profiles" & varsuffix).Rows.Count, Range("profileListStart" & varsuffix).Column)
                If primaryLogin Then
                    Call getAccountDataOuter(profileListStart, "PRIMARY", True)
                    If dataSource = "GA" Then
                        pctDone = pctDone + 1
                        Call updateProgress(pctDone, "Fetching segments for " & usernameDisp & "...")
                        Call refreshSegmentList
                    End If
                    primaryLogin = False
                Else
                    Call getAccountDataOuter(profileListStart, "SECONDARY", True)
                End If
            End If
        Next rivi
    End With
    Call hideProgressBox

End Sub


Sub getProfilesForLogin(authToken As String)
    On Error Resume Next
    Dim profileListStart As Range

    Call setDatasourceVariables

    Set profileListStart = configsheet.Cells(Range("profileListStart" & varsuffix).row + Range("profiles" & varsuffix).Rows.Count, Range("profileListStart" & varsuffix).Column)
    Call getAccountDataOuter(profileListStart, "SECONDARY")

End Sub

Public Function trimEM(str As String) As String
    If Left(str, 3) = "em$" Then
        trimEM = Right(str, Len(str) - 3)
    Else
        trimEM = str
    End If
End Function

Sub launchLoginStatusBox(Optional setNum As Integer = 1)
    On Error Resume Next
    Application.ScreenUpdating = False
    If debugMode = True Then On Error GoTo 0

    Call checkOperatingSystem

    Call setDatasourceVariables

    Call updateLoginStatus

    Call showLoginSetInLoginStatusBox(setNum)
    Call loginStatusBox.setLoginSetNum(setNum)
    If usingMacOSX = False Then loginStatusBox.Show

    ' On Error Resume Next
    ' If debugMode = True Then On Error GoTo 0

End Sub

Sub showLoginSetInLoginStatusBox(Optional setNum As Integer = 1)
    On Error Resume Next

    If debugMode = True Then On Error GoTo 0
    Dim rivi As Long
    Dim vrivi As Long

    Dim loginNum As Long
    loginNum = 0
    Dim firstLogin As Integer
    Dim shownLoginNum As Integer
    shownLoginNum = 0
    firstLogin = 12 * (setNum - 1) + 1
    If setNum > 1 Then
        loginStatusBox.prevPageB.Visible = True
    Else
        loginStatusBox.prevPageB.Visible = False
    End If
    loginStatusBox.nextPageB.Visible = False

    With Sheets("logins")
        vrivi = vikarivi(.Cells(1, loginInfoCol))
        For rivi = 1 To vrivi
            email = trimEM(.Cells(rivi, loginInfoCol).value)
            usernameDisp = .Cells(rivi, loginInfoCol + 4).value
            If email <> vbNullString Then
                If usernameDisp = vbNullString Then usernameDisp = email
                loginNum = loginNum + 1
                If loginNum >= firstLogin Then
                    If shownLoginNum < 12 Then
                        shownLoginNum = shownLoginNum + 1
                        licenseType = .Cells(rivi, loginInfoCol + 1).value
                        licenseDaysLeft = .Cells(rivi, loginInfoCol + 2).value
                        With loginStatusBox
                            With .Controls("un" & shownLoginNum)
                                .Caption = usernameDisp
                                .Visible = True
                            End With
                            With .Controls("lt" & shownLoginNum)
                                .Caption = licenseType
                                .Visible = True
                            End With
                            With .Controls("ldl" & shownLoginNum)
                                .Caption = licenseDaysLeft
                                .Visible = True
                            End With
                            .Controls("logout" & shownLoginNum).Visible = True
                        End With
                    Else
                        loginStatusBox.nextPageB.Visible = True
                        Exit For
                    End If
                End If
            End If
        Next rivi
    End With
    If shownLoginNum < 12 Then
        With loginStatusBox
            For rivi = shownLoginNum + 1 To 12
                .Controls("un" & rivi).Visible = False
                .Controls("lt" & rivi).Visible = False
                .Controls("ldl" & rivi).Visible = False
                .Controls("logout" & rivi).Visible = False
            Next rivi
        End With
    End If

End Sub



Sub storeLoginToSheet()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Dim rivi As Long

    Call setDatasourceVariables

    If usernameDisp = vbNullString Then usernameDisp = email

    With Sheets("logins")
        rivi = vikarivi(.Cells(1, loginInfoCol))
        If rivi = 1 And .Cells(1, loginInfoCol).value = vbNullString Then
            rivi = 1
        Else
            rivi = rivi + 1
        End If
        .Cells(rivi, loginInfoCol).value = "em$" & email
        If demoVersion = True Then
            .Cells(rivi, loginInfoCol + 1).value = "Trial"
        Else
            .Cells(rivi, loginInfoCol + 1).value = "Full"
        End If
        .Cells(rivi, loginInfoCol + 2).value = licenseDaysLeft
        .Cells(rivi, loginInfoCol + 3).value = password
        .Cells(rivi, loginInfoCol + 4).value = usernameDisp
    End With
End Sub

Sub updateLoginStatus()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Dim rivi As Long
    Dim vrivi As Long
    Dim email As String
    Dim emailsCombined As String


    Call setDatasourceVariables

    With Sheets("logins")
        vrivi = vikarivi(.Cells(1, loginInfoCol))
        For rivi = 1 To vrivi
            email = trimEM(.Cells(rivi, loginInfoCol).value)
            emailsCombined = emailsCombined & rscL1 & email
        Next rivi

        Call checkE(emailsCombined, dataSource, , True)

        For rivi = 1 To UBound(tempArr)

            licenseStatus = tempArr(rivi, 1)
            demoStatus = tempArr(rivi, 2)
            licenseDaysLeft = tempArr(rivi, 3)
            email = tempArr(rivi, 4)
            If licenseStatus = "VALID" Then
                demoVersion = False
            Else
                demoVersion = True
            End If

            If email <> vbNullString Then

                If licenseStatus = "INVALID" Then
                    .Cells(rivi, loginInfoCol + 1).value = "Full"
                Else
                    If demoVersion = True Then
                        .Cells(rivi, loginInfoCol + 1).value = "Trial"
                    Else
                        .Cells(rivi, loginInfoCol + 1).value = "Full"
                    End If
                End If
                If licenseDaysLeft <= 0 Or licenseDaysLeft = vbNullString Then
                    .Cells(rivi, loginInfoCol + 2).value = "EXPIRED"
                    Call markOneAccountAsExpired(email)
                Else
                    .Cells(rivi, loginInfoCol + 2).value = licenseDaysLeft
                    Call markOneAccountAsActive(email)
                End If
            End If
        Next rivi
        Erase tempArr
    End With
End Sub

Sub markOneAccountAsExpired(email As String)
    Application.ScreenUpdating = False
    On Error Resume Next

    Dim rivi As Long
    Call unprotectSheets

    Call setDatasourceVariables
    If usernameDisp = vbNullString Then usernameDisp = email
    With Range("profiles" & varsuffix)
        For rivi = 1 To .Rows.Count
            If trimEM(.Cells(rivi, 2).value) = email Then
                With configsheet.Shapes("_CB" & rivi)
                    .TextFrame.Characters.Text = "LICENSE EXPIRED: " & usernameDisp
                    .Fill.Visible = True
                    .Fill.Transparency = 0.15
                    .Fill.ForeColor.RGB = RGB(217, 217, 217)
                End With
            End If
        Next rivi
    End With
End Sub
Sub markOneAccountAsActive(email As String)
    Application.ScreenUpdating = False
    On Error Resume Next
    Dim rivi As Long
    Call unprotectSheets

    Call setDatasourceVariables

    With Range("profiles" & varsuffix)
        For rivi = 1 To .Rows.Count
            If trimEM(.Cells(rivi, 2).value) = email Then
                With configsheet.Shapes("_CB" & rivi)
                    .TextFrame.Characters.Text = vbNullString
                    .Fill.Visible = False
                End With
            End If
        Next rivi
    End With
End Sub

Sub logoutOneAccount(email As String)
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    If dataSource = "TW" Then Exit Sub
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim loginCount As Long
    Dim valuesArr As Variant

    Call setDatasourceVariables

    loginType = "SECONDARY"

    Dim rivi As Long
    Dim vrivi As Long
    With Sheets("logins")
        rivi = findRowWithValue(loginInfoCol, "em$" & email, 1, Sheets("logins"), 1)
        If rivi > 0 Then
            .Cells(rivi, loginInfoCol).Resize(1, 6).Clear
        End If
    End With


    With Sheets("tokens")
        vrivi = vikarivi(.Cells(1, loginInfoCol + 2))
        valuesArr = .Range(.Cells(1, loginInfoCol + 2), .Cells(vrivi, loginInfoCol + 2)).value
        For rivi = vrivi To 1 Step -1
            If trimEM(CStr(valuesArr(rivi, 1))) = CStr(email) Then
                .Cells(rivi, loginInfoCol).Resize(1, 6).Clear
            End If
        Next rivi
    End With
    valuesArr = ""

    Dim profileListStartRow As Long
    Dim profileListStartColumn As Long
    profileListStartRow = Range("profileListStart" & varsuffix).row
    profileListStartColumn = Range("profileListStart" & varsuffix).Column
    Dim firstRowToDelete As Integer
    firstRowToDelete = -1
    With Range("profiles" & varsuffix)
        valuesArr = .value
        For rivi = .Rows.Count To 1 Step -1
            If CStr(parseUserID(CStr(valuesArr(rivi, 2)))) = CStr(email) Then
                If firstRowToDelete = -1 Then
                    firstRowToDelete = rivi
                End If
            ElseIf firstRowToDelete <> -1 Then

                .Rows(rivi + 1 & ":" & firstRowToDelete).Delete xlShiftUp
                firstRowToDelete = -1

            End If
        Next rivi
        If firstRowToDelete <> -1 Then
            .Rows(1 & ":" & firstRowToDelete).Delete xlShiftUp
        End If
    End With
    valuesArr = ""
    With configsheet
        .Cells(profileListStartRow, profileListStartColumn).Name = "profileListStart" & varsuffix
        If Range("profileliststart" & varsuffix).value = vbNullString Then
            .Cells(profileListStartRow, profileListStartColumn - 2).Name = "profileSelections" & varsuffix
            .Cells(profileListStartRow, profileListStartColumn - 2).Resize(1, 5).Name = "profiles" & varsuffix
        End If
    End With

    Call addProfileSelectionCheckBoxes

    loginCount = Application.CountA(Sheets("logins").Columns(ColumnLetter(loginInfoCol)))

    If loginCount = 0 Then
        Call logout
    ElseIf loginCount = 1 Then
        Call setSingleAccountFormatting
    Else
        Call setMultiAccountFormatting
    End If

    Application.EnableEvents = True
End Sub


Sub storeTokenToSheet(ByVal profID As Variant, ByVal token As String, ByVal email As String, Optional profName As String, Optional accountName As String)
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Dim rivi As Long
    Dim i As Long
    Dim valuesArr As Variant
    Dim valuesRng As Range
    Dim storeNames As Boolean
    If profName <> vbNullString Then
        storeNames = True
        If accountName = vbNullString Then accountName = profName
    End If

    If token = "" Then
        If debugMode = True Then Debug.Print "Error: trying to insert empty token"
        Exit Sub
    End If

    Call setDatasourceVariables

    With Sheets("tokens")
        If profID = "ALL" Then
            Set valuesRng = .Range(.Cells(1, loginInfoCol), .Cells(vikarivi(.Cells(1, loginInfoCol)), loginInfoCol + 5))
            valuesArr = valuesRng.value
            For rivi = 1 To UBound(valuesArr)
                If trimEM(CStr(valuesArr(rivi, 3))) = email Then
                    valuesArr(rivi, 2) = token
                    If Len(token) > 200 Then
                        valuesArr(rivi, 5) = Left(token, 100) & Right(token, 100)
                    Else
                        valuesArr(rivi, 5) = token
                    End If
                    valuesArr(rivi, 4) = Now
                End If
            Next rivi
            valuesRng.value = valuesArr
        Else
            rivi = 0
            rivi = findRowWithValue(loginInfoCol, "id" & profID, 1, Sheets("tokens"), 1)
            If rivi = 0 Then
                rivi = vikarivi(.Cells(1, loginInfoCol)) + 1
            ElseIf trimEM(.Cells(rivi, loginInfoCol + 2).value) <> email Then
                For i = 1 To 5
                    rivi = findRowWithValue(loginInfoCol, "id" & profID, rivi, Sheets("tokens"), rivi + 1)
                    If rivi = 0 Then
                        rivi = vikarivi(.Cells(1, loginInfoCol)) + 1
                        Exit For
                    End If
                    If trimEM(.Cells(rivi, loginInfoCol + 2).value) <> email Then Exit For
                Next i
            End If
            .Cells(rivi, loginInfoCol).value = "id" & profID
            .Cells(rivi, loginInfoCol + 1).value = token
            .Cells(rivi, loginInfoCol + 2).value = "em$" & email
            .Cells(rivi, loginInfoCol + 3).value = Now
            If Len(token) > 200 Then
                .Cells(rivi, loginInfoCol + 4).value = Left(token, 100) & Right(token, 100)
            Else
                .Cells(rivi, loginInfoCol + 4).value = token
            End If
            If storeNames Then .Cells(rivi, loginInfoCol + 5).value = profName & "%%%" & accountName
        End If
    End With
End Sub

Sub storeEmailCheckedDateToSheet(ByVal email As String)
    On Error Resume Next
    Dim rivi As Long
    Dim vrivi As Long
    Dim valuesArr As Variant
    Dim valuesRng As Range

    Call setDatasourceVariables

    With Sheets("tokens")
        Set valuesRng = .Range(.Cells(1, loginInfoCol), .Cells(vikarivi(.Cells(1, loginInfoCol)), loginInfoCol + 3))
        valuesArr = valuesRng.value

        For rivi = 1 To UBound(valuesArr)
            If trimEM(CStr(valuesArr(rivi, 3))) = email Then valuesArr(rivi, 4) = Now
        Next rivi
        valuesRng.value = valuesArr
    End With
End Sub
Public Function getTokenFromSheet(ByVal profID As Variant) As String
    On Error Resume Next
    Dim rivi As Long
    Dim rivi2 As Long
    Dim timeLimit As Long
    Call setDatasourceVariables

    With Sheets("tokens")
        rivi = findRowWithValue(loginInfoCol, "id" & profID, 1, Sheets("tokens"), 1)
        If rivi = 0 Then
            'no token found for this profid, get first token
            For rivi2 = 1 To vikarivi(.Cells(1, loginInfoCol + 1))
                If .Cells(rivi2, loginInfoCol + 1).value <> vbNullString Then
                    rivi = rivi2
                    Exit For
                End If
            Next rivi2
        End If

        If rivi = 0 Then
            authToken = ""
        Else

            authToken = .Cells(rivi, loginInfoCol + 1).value
            email = trimEM(.Cells(rivi, loginInfoCol + 2).value)
            emailLastCheckedOK = .Cells(rivi, loginInfoCol + 3).value
            If dataSource = "GA" Or dataSource = "AW" Or dataSource = "YT" Or dataSource = "GW" Then
                timeLimit = 40 * 60
                If (Now - .Cells(rivi, loginInfoCol + 3).value) * 86400 > timeLimit Then
                    Debug.Print "Token is old (" & Round((Now - .Cells(rivi, loginInfoCol + 3).value) * 86400) & " seconds), refreshing for all profiles of this user (" & email & ")"
                    authToken = refreshToken(authToken)
                    Call storeTokenToSheet("ALL", authToken, email)
                End If
            End If
        End If
    End With

    getTokenFromSheet = authToken
End Function


Public Function getEmailForToken(ByVal token As String, Optional getPW As Boolean = False) As String
    On Error Resume Next
    Dim rivi As Long

    Call setDatasourceVariables

    With Sheets("tokens")
        If Len(token) > 200 Then
            rivi = findRowWithValue(loginInfoCol + 4, Left(token, 100) & Right(token, 100), 1, Sheets("tokens"), 1)
        Else
            rivi = findRowWithValue(loginInfoCol + 1, token, 1, Sheets("tokens"), 1)
        End If
        If rivi = 0 Then
            getEmailForToken = ""
        Else
            email = trimEM(.Cells(rivi, loginInfoCol + 2).value)
            getEmailForToken = email
            emailLastCheckedOK = .Cells(rivi, loginInfoCol + 3).value
        End If
    End With

    If getPW = True Then
        With Sheets("logins")
            rivi = findRowWithValue(loginInfoCol, "em$" & email, 1, Sheets("logins"), 1)
            If rivi > 0 Then
                password = .Cells(rivi, loginInfoCol + 3).value
            End If
        End With
    End If
End Function


Public Function getTokenForEmail(ByVal emailLoc As String) As String
    On Error Resume Next
    Dim rivi As Long

    Call setDatasourceVariables

    If Left(emailLoc, 2) <> "em$" Then emailLoc = "em$" & emailLoc

    With Sheets("tokens")
        rivi = findRowWithValue(loginInfoCol + 2, emailLoc, 1, Sheets("tokens"), 1)
        If rivi = 0 Then
            getTokenForEmail = ""
        Else
            getTokenForEmail = .Cells(rivi, loginInfoCol + 1).value
        End If
    End With

End Function

Public Function getFirstLoginEmail() As String
    On Error Resume Next
    Dim rivi As Long
    Dim vrivi As Long
    Call setDatasourceVariables

    With Sheets("logins")
        vrivi = vikarivi(.Cells(1, loginInfoCol))
        For rivi = 1 To vrivi
            If .Cells(rivi, loginInfoCol).value <> vbNullString Then
                getFirstLoginEmail = trimEM(.Cells(rivi, loginInfoCol).value)
                licenseDaysLeft = .Cells(rivi, loginInfoCol + 2).value
                licenseType = .Cells(rivi, loginInfoCol + 1).value
                usernameDisp = .Cells(rivi, loginInfoCol + 4).value
                Exit Function
            End If
        Next rivi
    End With
    getFirstLoginEmail = ""
End Function


Public Function getPWforEmail(ByVal email As String) As String
    On Error Resume Next
    Dim rivi As Long

    Call setDatasourceVariables

    With Sheets("logins")
        rivi = findRowWithValue(loginInfoCol, "em$" & email, 1, Sheets("logins"), 1)
        If rivi > 0 Then
            password = .Cells(rivi, loginInfoCol + 3).value
            getPWforEmail = .Cells(rivi, loginInfoCol + 3).value
        Else
            getPWforEmail = vbNullString
        End If
    End With
End Function


Sub addProfileSelectionCheckBoxes(Optional ByVal oldProfilesCount As Long = 0)

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim profileCount As Long
    Dim profileListStart As Range
    Dim buttonWidth As Double
    Dim solu As Range
    Dim cBox As Object
    Set profileListStart = Range("profileListStart" & varsuffix)
    profileCount = Range("profiles" & varsuffix).Rows.Count
    Call deleteProfileSelectionCBs

    buttonWidth = profileListStart.Offset(, -1).Resize(1, 4).Width


    If profileCount > 2000 Then
        MsgBox "Your " & referToProfilesAsSing & " list is so long (" & profileCount & " rows) that for performance reasons, it won't be possible to select a " & referToProfilesAsSing & " to a report by clicking it. Instead, to pick a " & referToProfilesAsSing & ", type any value to the blank cell to the left of the " & referToProfilesAsSing & " information", , "Note about long " & referToProfilesAsSing & " list"
    Else
        With configsheet
            For profNum = 1 To profileCount
                'CHECKBOXES
                Set solu = profileListStart.Offset(profNum - 1, -1)
                Set cBox = .Shapes.AddTextbox(1, solu.Left, solu.Top, buttonWidth, solu.Height)
                With cBox
                    .Fill.Visible = False
                    .Line.Visible = False
                    .Top = solu.Top
                    .TextFrame.Characters.Text = ""
                    .OnAction = "updateProfileSelections" & dataSource & "button"
                    .Left = solu.Left
                    .Name = "_CB" & profNum    '+ oldProfilesCount
                    .Width = buttonWidth
                    .Height = solu.Height
                End With
                DoEvents
                If profNum Mod 200 = 1 Then
                    Call updateProgressIterationBoxes
                End If
            Next profNum
        End With

        Call updateProgressIterationBoxes("EXITLOOP")
    End If

End Sub



Public Function getGoalName(profIDloc As Variant, goalIDloc As Variant) As String
    Dim rivi As Long
    Dim resultStr As String
    goalIDloc = CStr(goalIDloc)
    profIDloc = CStr(profIDloc)
    resultStr = vbNullString
    For rivi = 1 To UBound(goalsArr, 1)
        If goalsArr(rivi, 1) = profIDloc And CStr(goalsArr(rivi, 2)) = goalIDloc Then
            resultStr = goalsArr(rivi, 3)
            Exit For
        End If
    Next rivi
    getGoalName = resultStr
End Function



