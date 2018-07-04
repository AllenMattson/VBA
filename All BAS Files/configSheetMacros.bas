Attribute VB_Name = "configSheetMacros"
Option Private Module
Option Explicit


Sub protectSheets()
    On Error Resume Next
    Dim sh As Worksheet
    Set sh = ActiveSheet
    Application.ScreenUpdating = False
    Modules.Protect userinterfaceonly:=True
    Analytics.Protect userinterfaceonly:=True, AllowFiltering:=True
    AdWords.Protect userinterfaceonly:=True, AllowFiltering:=True
    BingAds.Protect userinterfaceonly:=True, AllowFiltering:=True
    YouTube.Protect userinterfaceonly:=True, AllowFiltering:=True
    Facebook.Protect userinterfaceonly:=True, AllowFiltering:=True
    Twitter.Protect userinterfaceonly:=True, AllowFiltering:=True
    Webmaster.Protect userinterfaceonly:=True, AllowFiltering:=True
    Stripe.Protect userinterfaceonly:=True, AllowFiltering:=True
    FacebookAds.Protect userinterfaceonly:=True, AllowFiltering:=True
    MailChimp.Protect userinterfaceonly:=True, AllowFiltering:=True
    TwitterAds.Protect userinterfaceonly:=True, AllowFiltering:=True
    sh.Select
    sheetProtectionRemoved = False
End Sub

Sub unprotectSheets()
    On Error Resume Next
    If Not sheetProtectionRemoved Then
        Application.ScreenUpdating = False
        Dim sh As Worksheet
        Set sh = ActiveSheet
        Analytics.Unprotect
        AdWords.Unprotect
        BingAds.Unprotect
        Modules.Unprotect
        YouTube.Unprotect
        Facebook.Unprotect
        Twitter.Unprotect
        Webmaster.Unprotect
        Stripe.Unprotect
        FacebookAds.Unprotect
        MailChimp.Unprotect
         TwitterAds.Unprotect
        sh.Select
        sheetProtectionRemoved = True
    End If
End Sub

Sub checkReportFormattingOptionsVisibility(Optional shname As String)
    On Error Resume Next
    Dim sh As Worksheet
    If shname = vbNullString Then
        Set sh = ActiveSheet
    Else
        Set sh = Sheets(shname)
    End If
    With sh
        If .OptionButtons("formattedReportOB").value = 1 Then
            .CheckBoxes("createChartsCB").Enabled = True
            .DropDowns("condFormDropDown").Enabled = True
            .DropDowns("groupingDD").Enabled = True
            .Shapes("condFormLabel").TextFrame.Characters.Font.ColorIndex = 1
            .Shapes("createChartsLabel").TextFrame.Characters.Font.ColorIndex = 1
            .Shapes("groupingLabel").TextFrame.Characters.Font.ColorIndex = 1
            If .Name = "YouTube" Or .Name = "Flickr" Then
                .Shapes("videoLinksLabel").TextFrame.Characters.Font.ColorIndex = 1
                .CheckBoxes("videoLinksCB").Enabled = True
            End If
        Else
            .CheckBoxes("createChartsCB").Enabled = False
            .DropDowns("condFormDropDown").Enabled = False
            .DropDowns("groupingDD").Enabled = False
            .Shapes("condFormLabel").TextFrame.Characters.Font.ColorIndex = 15
            .Shapes("createChartsLabel").TextFrame.Characters.Font.ColorIndex = 15
            .Shapes("groupingLabel").TextFrame.Characters.Font.ColorIndex = 15
            If .Name = "YouTube" Or .Name = "Flickr" Then
                .Shapes("videoLinksLabel").TextFrame.Characters.Font.ColorIndex = 15
                .CheckBoxes("videoLinksCB").Enabled = False
            End If
        End If
    End With
End Sub
Sub rowLimitDDchange()
    On Error Resume Next
    Dim inputValue As Long
    Dim rivi As Integer
    rivi = Application.Match("Custom", Range("rowLimitDDvalues"), 0)
    With Analytics
        .Unprotect
        If Sheets(.Name).Shapes("rowLimitDD").ControlFormat.value = rivi Then
            inputValue = CLng(InputBox("How many rows should be fetched per profile?", "Number of rows to fetch"))
            If inputValue <= 0 Then
                inputValue = 10
            ElseIf inputValue > 1000000 Then
                inputValue = 1000000
            End If
            Range("rowLimitDDvalues").Cells(rivi).Offset(, 1).value = inputValue
            With .Shapes("customRowLimit")
                .TextFrame.Characters.Text = inputValue & " rows"
                .Visible = True
            End With
        Else
            .Shapes("customRowLimit").Visible = False
        End If
        .Protect userinterfaceonly:=True
    End With
End Sub



Sub showMacroInstructions()
    On Error Resume Next
    Call unprotectSheets
    With Modules
        .Shapes("macroBox").Visible = True
        .Shapes("macroMessage").Visible = True
        .Shapes("macroMessage2").Visible = True
        .Shapes("macroInstructionsButton").Visible = True
    End With
End Sub
Sub hideMacroInstructions()
    On Error Resume Next
    Call unprotectSheets
    With Modules
        .Shapes("macroBox").Visible = False
        .Shapes("macroMessage").Visible = False
        .Shapes("macroMessage2").Visible = False
        .Shapes("macroInstructionsButton").Visible = False
    End With
End Sub
Sub showAutomationButtons()
    On Error Resume Next
    With Modules
        .Shapes("refreshButton").Visible = True
        .Shapes("exportButton").Visible = True
        .Shapes("copyButton").Visible = True
        .Shapes("deleteAllReportsButton").Visible = True
    End With
End Sub
Sub hideAutomationButtons()
    On Error Resume Next
    With Modules
        .Shapes("refreshButton").Visible = False
        .Shapes("exportButton").Visible = False
        .Shapes("copyButton").Visible = False
        .Shapes("deleteAllReportsButton").Visible = False
    End With
End Sub

Sub selectStartDate()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim dateBeforeChange As Date

    Call setDatasourceVariables

    Application.ScreenUpdating = False
    Call unprotectSheets
    If IsDate(Range("startDate" & varsuffix).value) Then
        selectedDate = Range("startDate" & varsuffix).value
    Else
        selectedDate = Now
    End If
    dateBeforeChange = selectedDate
    selectDatesLabel = "Pick start date"
    CalendarFrm.Show
    Range("startDate" & varsuffix).value = DateSerial(Year(selectedDate), Month(selectedDate), Day(selectedDate))
    Sheets(configsheet.Name).Shapes("startDateDisp").TextFrame.Characters.Text = selectedDate
    If selectedDate <> dateBeforeChange Then
        Range(Sheets(configsheet.Name).Shapes("dateRangeTypeDD").ControlFormat.LinkedCell).value = Application.Match("custom", Range("dateRangeTypes"), 0)  'workaround for Mac Excel 2016 bug
        Sheets(configsheet.Name).Shapes("dateRangeTypeDD").ControlFormat.value = Application.Match("custom", Range("dateRangeTypes"), 0)
        Call dateRangeTypeChange
    End If
    Call comparisonDatesChange
    Call protectSheets

End Sub
Sub selectEndDate()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Dim dateBeforeChange As Date
    Call setDatasourceVariables

    Application.ScreenUpdating = False
    Call unprotectSheets
    If IsDate(Range("endDate" & varsuffix).value) Then
        selectedDate = Range("endDate" & varsuffix).value
    Else
        selectedDate = Now
    End If
    dateBeforeChange = selectedDate
    selectDatesLabel = "Pick end date"
    CalendarFrm.Show
    Range("endDate" & varsuffix).value = DateSerial(Year(selectedDate), Month(selectedDate), Day(selectedDate))
    '  Range("endDateDisp" & varsuffix).value = DateSerial(Year(selectedDate), Month(selectedDate), Day(selectedDate))
    Sheets(configsheet.Name).Shapes("endDateDisp").TextFrame.Characters.Text = selectedDate

    If selectedDate <> dateBeforeChange Then
        Range(Sheets(configsheet.Name).Shapes("dateRangeTypeDD").ControlFormat.LinkedCell).value = Application.Match("custom", Range("dateRangeTypes"), 0)    'workaround for Mac Excel 2016 bug
        Sheets(configsheet.Name).Shapes("dateRangeTypeDD").ControlFormat.value = Application.Match("custom", Range("dateRangeTypes"), 0)
        Call dateRangeTypeChange
    End If
    Call comparisonDatesChange
    Call protectSheets

End Sub



Sub dateRangeTypeChange()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Application.ScreenUpdating = False

    Call unprotectSheets
    Call setDatasourceVariables

    If dataSource = "GA" Then
        Range("dateRangeTypeGA").Calculate
        dateRangeType = LCase(Range("dateRangeTypeGA").value)
    Else
        Range("dateRangeType" & varsuffix).Calculate
        dateRangeType = LCase(Range("dateRangeType" & varsuffix).value)
    End If


    With configsheet

        If InStr(1, dateRangeType, "lastx") > 0 Then
            .Shapes("lastXbox" & varsuffix).Visible = True
            .Shapes("lastXlabel1" & varsuffix).Visible = True
            .Shapes("lastXlabel2" & varsuffix).Visible = True



            If IsNumeric(.Shapes("lastXbox" & varsuffix).TextFrame.Characters.Text) Then
                If (.Shapes("lastXbox" & varsuffix).TextFrame.Characters.Text > 0) Then
                    dateRangeType = Replace(dateRangeType, "lastx", "last" & val(.Shapes("lastXbox" & varsuffix).TextFrame.Characters.Text))
                End If
            End If

            If .CheckBoxes("includeCurrentCB").value = 1 Then dateRangeType = dateRangeType & "inc"

        Else
            .Shapes("lastXbox" & varsuffix).Visible = False
            .Shapes("lastXlabel1" & varsuffix).Visible = False
            .Shapes("lastXlabel2" & varsuffix).Visible = False
            .Shapes("includeCurrentLabel" & varsuffix).Visible = False
            .CheckBoxes("includeCurrentCB").Visible = False
        End If



        timePeriodForXdateRange = ""
        Call getDatesForDateRangeType(dateRangeType)
        If timePeriodForXdateRange <> "" Then
            If timePeriodForXdateRange = "weeksiso" Then
                .Shapes("lastXlabel2" & varsuffix).TextFrame.Characters.Text = "weeks"
            Else
                .Shapes("lastXlabel2" & varsuffix).TextFrame.Characters.Text = timePeriodForXdateRange
            End If
            With .Shapes("includeCurrentLabel" & varsuffix)
                Select Case timePeriodForXdateRange
                Case "days"
                    .TextFrame.Characters.Text = "Including today"
                Case "weeks", "weeksiso"
                    .TextFrame.Characters.Text = "Including this week"
                Case "months"
                    .TextFrame.Characters.Text = "Including this month"
                Case "years"
                    .TextFrame.Characters.Text = "Including this year"
                End Select
                .Visible = True
            End With
            .CheckBoxes("includeCurrentCB").Visible = True
        End If



        If dateRangeType <> "fixed" And dateRangeType <> "custom" Then
            Range("startDate" & varsuffix).value = DateSerial(Year(startDate), Month(startDate), Day(startDate))
            Range("endDate" & varsuffix).value = DateSerial(Year(endDate), Month(endDate), Day(endDate))
            .Shapes("startDateDisp").TextFrame.Characters.Text = startDate
            .Shapes("endDateDisp").TextFrame.Characters.Text = endDate

            '       Range("startDateDisp" & varsuffix).value = DateSerial(Year(startDate), Month(startDate), Day(startDate))
            '       Range("endDateDisp" & varsuffix).value = DateSerial(Year(endDate), Month(endDate), Day(endDate))
        End If
    End With
    Call protectSheets
End Sub

Sub changeDateRangeTypeX()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Application.ScreenUpdating = False
    Call unprotectSheets
    Call setDatasourceVariables
    Dim selectedX As Variant
    With configsheet
        selectedX = InputBox("Select number of " & .Shapes("lastXlabel2" & varsuffix).TextFrame.Characters.Text & " to fetch", "Select number of " & .Shapes("lastXlabel2" & varsuffix).TextFrame.Characters.Text)

        If selectedX = "" Then
            Exit Sub
        ElseIf Not IsNumeric(selectedX) Then
            MsgBox "Invalid value. Please input a number.", , "Invalid value"
        ElseIf val(selectedX) <= 0 Then
            MsgBox "Input a number greater than zero please.", , "Invalid value"
        Else
            .Shapes("lastXbox" & varsuffix).TextFrame.Characters.Text = val(selectedX)
            Call dateRangeTypeChange
        End If
    End With
    Call protectSheets
End Sub
Sub getDatesForDateRangeType(dateRangeType As String)
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim includeCurrent As Boolean
    Dim timePeriod As String

    Dim selectedX As Long

    Select Case dateRangeType
    Case "fixed", "custom"
        startDate = Range("startDate" & varsuffix).value
        endDate = Range("endDate" & varsuffix).value
    Case "today"
        startDate = Date
        endDate = Date
    Case "yesterday"
        startDate = Date - 1
        endDate = Date - 1
    Case "lastweeksunmon", "lastweek"
        startDate = sundayOfLastWeek(Date)
        endDate = startDate + 6
    Case "lastweekmonsun"
        startDate = mondayOfLastWeek(Date)
        endDate = startDate + 6
    Case "thismonth"
        startDate = DateSerial(Year(Date), Month(Date), 1)
        endDate = Date
    Case "lastmonth"
        endDate = DateSerial(Year(Date), Month(Date), 0)
        startDate = DateSerial(Year(endDate), Month(endDate), 1)
    Case "thisyear"
        startDate = DateSerial(Year(Date), 1, 1)
        endDate = Date
    Case "lastyear"
        endDate = DateSerial(Year(Date), 1, 0)
        startDate = DateSerial(Year(endDate), 1, 1)
    Case "lastyeartodate"
        endDate = Date
        startDate = DateSerial(Year(endDate) - 1, 1, 1)
    Case "last2yearstodate"
        endDate = Date
        startDate = DateSerial(Year(endDate) - 2, 1, 1)
    Case "last3yearstodate"
        endDate = Date
        startDate = DateSerial(Year(endDate) - 3, 1, 1)
    Case Else

        Dim Counter As Integer
        Dim numberFound As Boolean
        Dim xValueInStr As String
        numberFound = False
        For Counter = 1 To Len(dateRangeType)
            If IsNumeric(Mid(dateRangeType, Counter, 1)) Then
                numberFound = True
                xValueInStr = xValueInStr & Mid(dateRangeType, Counter, 1)
            ElseIf numberFound = False Then

            Else
                timePeriod = timePeriod & Mid(dateRangeType, Counter, 1)
            End If
        Next

        selectedX = val(xValueInStr)

        If InStr(1, timePeriod, "inc") > 0 Then
            includeCurrent = True
            timePeriod = Replace(timePeriod, "inc", "")
        Else
            includeCurrent = False
        End If

        If timePeriod = vbNullString Then timePeriod = "days"

        timePeriodForXdateRange = timePeriod

        Select Case timePeriod
        Case "days"
            If includeCurrent = True Then
                startDate = Date - selectedX + 1
                endDate = Date
            Else
                startDate = Date - selectedX
                endDate = Date - 1
            End If
        Case "weeks"
            If includeCurrent = True Then
                startDate = DateAdd("ww", -(selectedX) + 2, sundayOfLastWeek(Date))
                endDate = Date
            Else
                startDate = DateAdd("ww", -(selectedX) + 1, sundayOfLastWeek(Date))
                endDate = startDate + selectedX * 7 - 1
            End If
        Case "weeksiso"
            If includeCurrent = True Then
                startDate = mondayOfLastWeek(Date) + 7 - (selectedX - 1) * 7
                endDate = Date
            Else
                startDate = mondayOfLastWeek(Date) - (selectedX - 1) * 7
                endDate = startDate + selectedX * 7 - 1
            End If
        Case "months"
            If includeCurrent = True Then
                startDate = DateSerial(Year(DateAdd("m", -selectedX + 1, Date)), Month(DateAdd("m", -selectedX + 1, Date)), 1)
                endDate = Date
            Else
                startDate = DateSerial(Year(DateAdd("m", -selectedX, Date)), Month(DateAdd("m", -selectedX, Date)), 1)
                endDate = DateSerial(Year(Date), Month(Date), 0)
            End If
        Case "years"
            If includeCurrent = True Then
                endDate = Date
                startDate = DateSerial(Year(endDate) - selectedX + 1, 1, 1)
            Else
                endDate = DateSerial(Year(Date) - 1, 12, 31)
                startDate = DateSerial(Year(endDate) - selectedX + 1, 1, 1)
            End If
            'lastxinc
            'lastxweeks
            'lastxweeksinc
            'lastxmonths
            'lastxmonthsinc
            'lastxyears
            'lastxyearsinc

        End Select

    End Select


End Sub

Public Function sundayOfLastWeek(forDate As Date) As Date
    Dim curWeekDay As Integer
    curWeekDay = WeekDay(forDate)
    sundayOfLastWeek = forDate - curWeekDay - 6
End Function
Public Function mondayOfLastWeek(forDate As Date) As Date
    Dim curWeekDay As Integer
    curWeekDay = WeekDay(forDate) - 1
    If curWeekDay = 0 Then curWeekDay = 7
    mondayOfLastWeek = forDate - curWeekDay - 6
End Function

Public Function getDispNameForDateRangeType(dateRangeType As String) As String
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim includeCurrent As Boolean
    Dim timePeriod As String

    Dim selectedX As Long


    Select Case dateRangeType
    Case "fixed", "custom"
        getDispNameForDateRangeType = "Fixed dates"
    Case "today", "last1inc"
        getDispNameForDateRangeType = "Today"
    Case "yesterday", "last1"
        getDispNameForDateRangeType = "Yesterday"
    Case "last1weeksinc"
        getDispNameForDateRangeType = "This week"
    Case "lastweeksunmon", "lastweek", "last1weeks"
        getDispNameForDateRangeType = "Last week"
    Case "lastweekmonsun"
        getDispNameForDateRangeType = "Last week"
    Case "thismonth", "last1monthsinc"
        getDispNameForDateRangeType = "This month to date"
    Case "lastmonth", "last1months"
        getDispNameForDateRangeType = "Last month"
    Case "thisyear", "last1yearsinc"
        getDispNameForDateRangeType = "This year to date"
    Case "lastyear", "last1years"
        getDispNameForDateRangeType = "Last year"
    Case "lastyeartodate", "last2yearsinc"
        getDispNameForDateRangeType = "Last year & this year to date"
    Case "last2yearstodate", "last3yearsinc"
        getDispNameForDateRangeType = "Last 2 years & this year to date"
    Case "last3yearstodate", "last4yearsinc"
        getDispNameForDateRangeType = "Last 3 years & this year to date"
    Case Else

        Dim Counter As Integer
        Dim numberFound As Boolean
        Dim xValueInStr As String
        numberFound = False
        For Counter = 1 To Len(dateRangeType)
            If IsNumeric(Mid(dateRangeType, Counter, 1)) Then
                numberFound = True
                xValueInStr = xValueInStr & Mid(dateRangeType, Counter, 1)
            ElseIf numberFound = False Then

            Else
                timePeriod = timePeriod & Mid(dateRangeType, Counter, 1)
            End If
        Next

        selectedX = val(xValueInStr)

        If InStr(1, timePeriod, "inc") > 0 Then
            includeCurrent = True
            timePeriod = Replace(timePeriod, "inc", "")
        Else
            includeCurrent = False
        End If
        If timePeriod = vbNullString Then timePeriod = "days"
        If timePeriod = "weeksiso" Then timePeriod = "weeks"
        getDispNameForDateRangeType = "Last " & selectedX & " " & timePeriod

        If includeCurrent = True Then
            Select Case timePeriod
            Case "days"
                getDispNameForDateRangeType = getDispNameForDateRangeType & " (including today)"
            Case "weeks", "weeksiso"
                getDispNameForDateRangeType = getDispNameForDateRangeType & " (including this week)"
            Case "months"
                getDispNameForDateRangeType = getDispNameForDateRangeType & " (including this month)"
            Case "years"
                getDispNameForDateRangeType = getDispNameForDateRangeType & " (including this year)"
            End Select
        End If

    End Select
End Function




Sub showProxySettings()

    On Error Resume Next

    Application.ScreenUpdating = False
    Sheets("proxysettings").Visible = xlSheetVisible

End Sub
Sub hideProxySettings()

    On Error Resume Next
    Application.ScreenUpdating = False
    Sheets("proxysettings").Visible = xlSheetHidden

End Sub


Sub setSingleAccountFormatting()
    On Error Resume Next

    Call setDatasourceVariables

    email = getFirstLoginEmail()

    If usernameDisp = vbNullString Then email = usernameDisp

    With configsheet
        .Shapes("logoutButton" & varsuffix).Visible = True
        .Shapes("logoutButtonMultiAccount" & varsuffix).Visible = False
        .Shapes("addLoginButton" & varsuffix).Visible = True
        .Shapes("addLoginButtonNote1" & varsuffix).Visible = True
        .Shapes("addLoginButtonNote2" & varsuffix).Visible = True
        .Shapes("manageLoginsButton" & varsuffix).Visible = False
        .Shapes("licenseNote" & varsuffix).Visible = True
        .Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Logged in with account " & usernameDisp
        If demoVersion Then
            .Shapes("licenseNote" & varsuffix).TextFrame.Characters.Text = "Trial days left: " & licenseDaysLeft
        Else
            .Shapes("licenseNote" & varsuffix).TextFrame.Characters.Text = "License days left: " & licenseDaysLeft
        End If
    End With
    With Modules
        .Shapes("logoutButton" & varsuffix).Visible = True
        ' .Shapes("logoutButtonMultiAccount").Visible = False
        .Shapes("addLoginButton" & varsuffix).Visible = True
        .Shapes("addLoginButtonNote1" & varsuffix).Visible = True
        .Shapes("addLoginButtonNote2" & varsuffix).Visible = True
        .Shapes("manageLoginsButton" & varsuffix).Visible = False
        .Shapes("licenseNote" & varsuffix).Visible = True
        .Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Logged in with account " & usernameDisp
        If demoVersion Then
            .Shapes("licenseNote" & varsuffix).TextFrame.Characters.Text = "Trial days left: " & licenseDaysLeft
        Else
            .Shapes("licenseNote" & varsuffix).TextFrame.Characters.Text = "License days left: " & licenseDaysLeft
        End If
    End With

End Sub
Sub setMultiAccountFormatting()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Call unprotectSheets

    Call setDatasourceVariables

    Dim loginCount As Long
    loginCount = Application.CountA(Sheets("logins").Columns(ColumnLetter(loginInfoCol)))
    With configsheet
        .Shapes("logoutButton" & varsuffix).Visible = False
        .Shapes("logoutButtonMultiAccount" & varsuffix).Visible = True
        .Shapes("addLoginButton" & varsuffix).Visible = False
        .Shapes("addLoginButtonNote1" & varsuffix).Visible = False
        .Shapes("addLoginButtonNote2" & varsuffix).Visible = False
        .Shapes("manageLoginsButton" & varsuffix).Visible = True
        .Shapes("licenseNote" & varsuffix).Visible = False
        .Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Logged in with " & loginCount & " accounts"
    End With
    With Modules
        .Shapes("logoutButton" & varsuffix).Visible = True
        '.Shapes("logoutButtonMultiAccount").Visible = True
        .Shapes("addLoginButton" & varsuffix).Visible = False
        .Shapes("addLoginButtonNote1" & varsuffix).Visible = False
        .Shapes("addLoginButtonNote2" & varsuffix).Visible = False
        .Shapes("manageLoginsButton" & varsuffix).Visible = True
        .Shapes("licenseNote" & varsuffix).Visible = False
        .Shapes("authStatusBox" & varsuffix).TextFrame.Characters.Text = "Logged in with " & loginCount & " accounts"
    End With
End Sub



Sub updateVisibilityOfDropdowns(dropDownCode As String, Optional clearSelections As Boolean = False)

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Application.ScreenUpdating = False
    Call unprotectSheets

    If debugMode = True Then Debug.Print "Running updateVisibilityOfDropdowns"

    Dim varsSheet As Worksheet

    Dim startCell As Range

    Dim itemCount As Long
    Dim itemNum As Long
    Dim foundDropdownWithNoSelection As Boolean
    Dim fieldNamesArr As Variant


    dataSource = UCase(Right(dropDownCode, 2))
    Call setDatasourceVariables
    Range("dimensionscalc" & varsuffix).Calculate

 

    Select Case dropDownCode
    Case "drdga"
        Set startCell = Range("dimension1name" & varsuffix)
        itemCount = 7
        Range("dimensionscalc").Calculate
    Case "drsdga"
        Set startCell = Range("segmDimension1name" & varsuffix)
        itemCount = 2
        Range("segmdimcalc").Calculate
    Case "drmga"
        Set startCell = Range("metric1name" & varsuffix)
        itemCount = 12
        Range("metricscalc").Calculate
    Case "drdaw"
        Set startCell = Range("dimension1name" & varsuffix).Offset(1, 0)
        itemCount = 6
        Range("dimensionscalcaw").Calculate
    Case "drsdaw"
        Set startCell = Range("segmDimension1name" & varsuffix)
        itemCount = 2
        Range("segmdimcalc").Calculate
    Case "drmaw"
        Set startCell = Range("metric1name" & varsuffix)
        itemCount = 12
        Range("metricscalcaw").Calculate
    Case "drtdaw"
        Exit Sub
        Set startCell = Range("dimension1name" & varsuffix)
        itemCount = 1
        Range("dimensionscalcaw").Calculate
    Case "drdac"
        Set startCell = Range("dimension1name" & varsuffix).Offset(1, 0)
        itemCount = 6
        Range("dimensionscalcac").Calculate
    Case "drsdac"
        Set startCell = Range("segmDimension1name" & varsuffix)
        itemCount = 2
        Range("segmdimcalc").Calculate
    Case "drmac"
        Set startCell = Range("metric1name" & varsuffix)
        itemCount = 12
        Range("metricscalcac").Calculate
    Case "drtdac"
        Exit Sub
        Set startCell = Range("dimension1name" & varsuffix)
        itemCount = 1
        Range("dimensionscalcac").Calculate
    Case "drsdfb"
        Set startCell = Range("segmDimension1name" & varsuffix)
        itemCount = 2
        Range("segmdimcalc").Calculate
    Case "drmfb"
        Set startCell = Range("metric1name" & varsuffix)
        itemCount = 12
        Range("metricscalcfb").Calculate
    Case "drdfb"
        Set startCell = Range("dimension1name" & varsuffix)
        itemCount = 10
        Range("dimensionscalcfb").Calculate
    Case "drsdyt"
        Set startCell = Range("segmDimension1name" & varsuffix)
        itemCount = 2
        Range("segmdimcalc").Calculate
    Case "drmyt"
        Set startCell = Range("metric1name" & varsuffix)
        itemCount = 12
        Range("metricscalcyt").Calculate
    Case "drdyt"
        Set startCell = Range("dimension1name" & varsuffix)
        itemCount = 5
        Range("dimensionscalcyt").Calculate
    Case "drsdgw"
        Set startCell = Range("segmDimension1name" & varsuffix)
        itemCount = 2
        Range("segmdimcalc").Calculate
    Case "drmgw"
        Set startCell = Range("metric1name" & varsuffix)
        itemCount = 5
        Range("metricscalcgw").Calculate
    Case "drdgw"
        Set startCell = Range("dimension1name" & varsuffix)
        itemCount = 5
        Range("dimensionscalcgw").Calculate
    Case "drsdfl"
        Set startCell = Range("segmDimension1name" & varsuffix)
        itemCount = 2
        Range("segmdimcalc").Calculate
    Case "drmfl"
        Set startCell = Range("metric1name" & varsuffix)
        itemCount = 12
        Range("metricscalcfl").Calculate
    Case "drdfl"
        Set startCell = Range("dimension1name" & varsuffix)
        itemCount = 3
        Range("dimensionscalcfl").Calculate
    Case "drsdst"
        Set startCell = Range("segmDimension1name" & varsuffix)
        itemCount = 2
        Range("segmdimcalc").Calculate
    Case "drmst"
        Set startCell = Range("metric1name" & varsuffix)
        itemCount = 10
        Range("metricscalcgw").Calculate
    Case "drdst"
        Set startCell = Range("dimension1name" & varsuffix)
        itemCount = 10
        Range("dimensionscalcgw").Calculate
    Case "drdfa"
        Set startCell = Range("dimension1name" & varsuffix)
        itemCount = 6
        Range("dimensionscalcfa").Calculate
    Case "drsdfa"
        Set startCell = Range("segmDimension1name" & varsuffix)
        itemCount = 2
        Range("segmdimcalc").Calculate
    Case "drmfa"
        Set startCell = Range("metric1name" & varsuffix)
        itemCount = 12
        Range("metricscalcfa").Calculate
    Case "drdmc"
        Set startCell = Range("dimension1name" & varsuffix)
        itemCount = 6
        Range("dimensionscalcmc").Calculate
    Case "drsdmc"
        Set startCell = Range("segmDimension1name" & varsuffix)
        itemCount = 2
        Range("segmdimcalc").Calculate
    Case "drmmc"
        Set startCell = Range("metric1name" & varsuffix)
        itemCount = 12
        Range("metricscalcmc").Calculate
    Case "drdta"
        Set startCell = Range("dimension1name" & varsuffix)
        itemCount = 10
        Range("dimensionscalcta").Calculate
    Case "drsdta"
        Set startCell = Range("segmDimension1name" & varsuffix)
        itemCount = 2
        Range("segmdimcalc").Calculate
    Case "drmta"
        Set startCell = Range("metric1name" & varsuffix)
        itemCount = 12
        Range("metricscalcta").Calculate
    End Select



    foundDropdownWithNoSelection = False

    For itemNum = 1 To itemCount
        If clearSelections Then
            Range(Sheets(configsheet.Name).Shapes(dropDownCode & "_" & itemNum).ControlFormat.LinkedCell).value = 1  'workaround for Mac Excel 2016 bug
            Sheets(configsheet.Name).Shapes(dropDownCode & "_" & itemNum).ControlFormat.value = 1   'need to use sheet name reference instead of object due to Mac Office bug
            If itemNum = 1 Then
                Sheets(configsheet.Name).Shapes(dropDownCode & "_" & itemNum).Visible = True
            Else
                Sheets(configsheet.Name).Shapes(dropDownCode & "_" & itemNum).Visible = False
            End If
        Else
            Sheets(configsheet.Name).Shapes(dropDownCode & "_" & itemNum).Visible = True
        End If
    Next itemNum

    'check whether to show SD category selection dd

    Range("segmDimension1name" & varsuffix).Calculate
    Range("segmDimension2name" & varsuffix).Calculate

    ' If Sheets(configsheet.Name).Shapes("drsd" & dataSource & "_1").ControlFormat.value <> 1 Or Sheets(configsheet.Name).Shapes("drsd" & dataSource & "_2").ControlFormat.value <> 1 Then
    If fieldNameIsOk(Range("segmDimension1name" & varsuffix).value) Or fieldNameIsOk(Range("segmDimension2name" & varsuffix).value) Then
        Sheets(configsheet.Name).Shapes("sdCategoriesDropdown").Visible = True
        Sheets(configsheet.Name).Shapes("sdCategoriesNote").Visible = True
    Else
        Sheets(configsheet.Name).Shapes("sdCategoriesDropdown").Visible = False
        Sheets(configsheet.Name).Shapes("sdCategoriesNote").Visible = False
    End If



    If clearSelections = False Then
        With varsSheetForDataSource
            For itemNum = 1 To itemCount
                If foundDropdownWithNoSelection = True Then
                    If fieldNameIsOk(.Cells(itemNum + startCell.row - 1, startCell.Column).value) = False Then
                        configsheet.Shapes(dropDownCode & "_" & itemNum).Visible = False
                    Else
                        configsheet.Shapes(dropDownCode & "_" & itemNum).Visible = True
                    End If
                End If

                If foundDropdownWithNoSelection = False And fieldNameIsOk(.Cells(itemNum + startCell.row - 1, startCell.Column).value) = False Then
                    foundDropdownWithNoSelection = True
                    configsheet.Shapes(dropDownCode & "_" & itemNum).Visible = True

                End If
            Next itemNum
        End With
    End If



    Call protectSheets

End Sub

Sub comparisonDatesChange()

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Application.ScreenUpdating = False
    Call unprotectSheets


    Call setDatasourceVariables

    Application.EnableEvents = False

    Dim startDate1 As Variant
    Dim endDate1 As Variant
    Dim comparisonType As String
    DoEvents
    Range("comparisonTypecurrent" & varsuffix).Calculate
    comparisonType = Range("comparisonTypecurrent" & varsuffix).value

    startDate1 = Range("startDate" & varsuffix).value
    endDate1 = Range("endDate" & varsuffix).value

    If comparisonType = "none" Then
        Range("startDateComparison" & varsuffix).ClearContents
        Range("endDateComparison" & varsuffix).ClearContents
        configsheet.DropDowns("comparisonValueTypeDD").Visible = False
    ElseIf comparisonType = "yearly" Then
        Range("startDateComparison" & varsuffix).value = DateSerial(Year(startDate1) - 1, Month(startDate1), Day(startDate1))
        Range("endDateComparison" & varsuffix).value = DateSerial(Year(endDate1) - 1, Month(endDate1), Day(endDate1))
        configsheet.DropDowns("comparisonValueTypeDD").Visible = True
    ElseIf comparisonType = "previous" Then
        Range("endDateComparison" & varsuffix).value = startDate1 - 1
        Range("startDateComparison" & varsuffix).value = Range("endDateComparison" & varsuffix).value - (endDate1 - startDate1)
        configsheet.DropDowns("comparisonValueTypeDD").Visible = True
    End If


    Application.EnableEvents = True

    Call protectSheets

End Sub



Sub hideSDselectionsYT()
    With YouTube
        .Shapes("sdimBox").Visible = False
        .Shapes("sdCategoriesNote").Visible = False
        .Shapes("drsdyt_1").Visible = False
        .Shapes("sdCategoriesDropdown").Visible = False
    End With
End Sub
Sub showSDselectionsYT()
    With YouTube
        .Shapes("sdimBox").Visible = True
        .Shapes("sdCategoriesNote").Visible = False
        .Shapes("drsdyt_1").Visible = True
        .Shapes("sdCategoriesDropdown").Visible = False
        Range("segmDimension1nameYT").Calculate
        If fieldNameIsOk(Range("segmDimension1nameYT").value) Then
            .Shapes("sdCategoriesDropdown").Visible = True
            .Shapes("sdCategoriesNote").Visible = True
        End If
    End With
End Sub
Sub hideFilterYT()
    With YouTube
        .Shapes("filterHeader").Visible = False
        .Shapes("filterButton").Visible = False
        ' .Shapes("filterButtonLogo").Visible = False
        .Shapes("filterNoteYT").Visible = False
    End With
    Call clearFiltersYT
End Sub
Sub showFilterYT()
    With YouTube
        .Shapes("filterHeader").Visible = True
        .Shapes("filterButton").Visible = True
        '        .Shapes("filterButtonLogo").Visible = True
        .Shapes("filterNoteYT").Visible = True
    End With
End Sub

Sub deleteProfileSelectionCBs()
    On Error Resume Next
    Call unprotectSheets
    Call setDatasourceVariables
    Dim cBox As Object

    With configsheet
        For Each cBox In .Shapes
            If Left(cBox.Name, 3) = "_CB" Then cBox.Delete
        Next
    End With

End Sub


Sub clearProfileSelections()
    On Error Resume Next

    Call setDatasourceVariables

    Application.ScreenUpdating = False
    Call unprotectSheets

    Range("profiles" & varsuffix).Columns(1).ClearContents
    If configsheet.AutoFilter.FilterMode Then
        calculationSetting = Application.Calculation
        Application.Calculation = xlManual
        Dim c As Range
        For Each c In Range("profiles" & varsuffix).Columns(1).Cells
            If c.EntireRow.Hidden Then
                c.value = ""
                With c.Resize(1, 5)
                    .Interior.ColorIndex = 2
                    .Font.ColorIndex = 1
                End With
            End If
        Next
        Application.Calculation = calculationSetting
    End If

    configsheet.Shapes("clearProfileSelectionsButton").Visible = False
    configsheet.Shapes("selectAllProfilesButton").Visible = True
    Call updateProfileSelections(True)
    Call protectSheets
End Sub

Sub selectAllProfiles()
    On Error Resume Next

    Call setDatasourceVariables

    Application.ScreenUpdating = False
    Call unprotectSheets
    Range("profiles" & varsuffix).Columns(1).value = "X"
    configsheet.Shapes("clearProfileSelectionsButton").Visible = True
    '  configsheet.Shapes("clearProfileSelectionsButtonIcon" & varsuffix).Visible = True
    configsheet.Shapes("selectAllProfilesButton").Visible = False
    Call updateProfileSelections(True)
    Call protectSheets
End Sub

Sub updateProfileSelectionsButton()

    On Error Resume Next
    Application.ScreenUpdating = False

    Call setDatasourceVariables

    Application.EnableEvents = False
    profNum = CInt(Right(Application.Caller, Len(Application.Caller) - 3))
    With Range("profilelistStart" & varsuffix).Offset(profNum - 1, -2)
        If .value = vbNullString Then
            .value = "X"
        Else
            .value = vbNullString
        End If
    End With

    Call updateProfileSelections(True)
    Call protectSheets

End Sub


Sub updateProfileSelections(Optional fromButton As Boolean = False)

    On Error Resume Next
    Application.ScreenUpdating = False

    Dim rivi As Long
    Dim loginCount As Long

    Call setDatasourceVariables

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    configsheet.Unprotect


    Dim vrivi As Long
    Dim profilesRange As Range
    Set profilesRange = Range("profiles" & varsuffix)

    Dim profRowsArr() As Variant

    profRowsArr = profilesRange.value

    loginCount = Application.CountA(Sheets("logins").Columns(ColumnLetter(loginInfoCol)))

    With profilesRange

        .Interior.ColorIndex = 2
        .Font.ColorIndex = 1

        If Not profilesRange.Worksheet.AutoFilter.FilterMode Then
            For rivi = 1 To profilesRange.Rows.Count
                If profRowsArr(rivi, 1) <> vbNullString Then
                    With .Rows(rivi)
                        .Interior.ColorIndex = 24
                        .Font.ColorIndex = 1
                    End With
                End If
            Next rivi
        Else
            calculationSetting = Application.Calculation
            Application.Calculation = xlManual
            For rivi = 1 To profilesRange.Rows.Count
                If profRowsArr(rivi, 1) <> vbNullString Then
                    With .Rows(rivi)
                        .Interior.ColorIndex = 24
                        .Font.ColorIndex = 1
                    End With
                Else
                    '                    With .Rows(rivi)
                    '                        If .Hidden Then
                    '                        .Interior.ColorIndex = 2
                    '                        .Font.ColorIndex = 1
                    '                        End If
                    '                    End With
                End If
            Next rivi
            Application.Calculation = calculationSetting
        End If
    End With

    Application.EnableEvents = True

End Sub

Sub hideSheet()
    ActiveSheet.Visible = xlSheetHidden
End Sub

