Attribute VB_Name = "queryManagement"
Option Private Module
Option Explicit

Sub saveQueryFromCQ()
    Sheets("vars").Cells(1, Range("parameterListStart").Column + 1).EntireColumn.Copy Range("savedQueryCol").EntireColumn
End Sub

Sub returnSavedQueryToCQ()
    Range("savedQueryCol").EntireColumn.Copy Range("parameterListStart").Offset(, 1).EntireColumn
End Sub

Sub setSampleQueryToCQ()
    Range("sampleQueryCol").EntireColumn.Copy Range("parameterListStart").Offset(, 1).EntireColumn
    Call getFromCurrentQuery(False)
End Sub

Sub clearQueryStorage()
    Sheets("querystorage").Columns("C:IV").ClearContents
End Sub

Sub removeSheet()

    On Error Resume Next
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Dim n As Name
    Dim sht As String
    Dim chartObj As Object

    Dim sh As Worksheet
    Set sh = ActiveSheet


    If isSheetAconfigSheet(sh.Name) = True Then Exit Sub


    Dim sheetID As String
    sheetID = Cells(1, 1).value
    sheetID = findRangeName(Cells(1, 1))

    ' Put in name of sheet where the range is located
    sht = sh.Name

    sh.Visible = xlSheetVeryHidden
    sh.Visible = xlSheetHidden

    'Excel 2010 freezes when deleting names
    'If CInt(left$(Application.Version, 2)) < 14 Then
    For Each n In ActiveWorkbook.Names
        If n.RefersToRange.Worksheet.Name = sht Then
            n.Delete
        End If
    Next n
    '    Else
    '        ' Application.DisplayAlerts = True
    '    End If

    Call deleteFromQueryStorageByID(sheetID)

    sh.Cells.Clear
    ' Application.ScreenUpdating = True
    ' ActiveWindow.FreezePanes = False
    '  Sheets("settings").Visible = xlSheetVisible
    ' Analytics.Select


    'Application.Wait (Now + TimeValue("00:00:01"))

    Sheets(sht).Delete
    ' sh.Delete

    '  Sheets("settings").Visible = xlSheetHidden

    Application.DisplayAlerts = True

End Sub

Public Function createNewSheetID() As String

    If debugMode = False Then On Error Resume Next

    Dim sheetID As String
    Randomize

    Do
        sheetID = "_SH" & Round(100000 * Rnd, 0)
        If Application.CountIf(Sheets("querystorage").Rows(Range("querySheetIDrow").row), sheetID) = 0 Then
            createNewSheetID = sheetID
            Exit Function
        End If
    Loop

End Function


Sub deleteNonCongfigSheets()
    On Error Resume Next
    Dim thisSheet As Worksheet
    Set thisSheet = ActiveSheet
    Application.ScreenUpdating = False
    Dim sh As Worksheet
    Application.DisplayAlerts = False
    For Each sh In ThisWorkbook.Worksheets
        If isSheetAconfigSheet(sh.Name) = False Then
            sh.Select
            Application.StatusBar = "Deleting sheet " & sh.Name
            Call removeSheet
        End If
    Next
    Application.DisplayAlerts = True
    Application.StatusBar = False
    thisSheet.Select
End Sub
Public Function isSheetAconfigSheet(sheetName As String) As Boolean
    sheetName = LCase$(sheetName)

    Select Case sheetName
    Case "analytics", "adwords", "twitterads", "settings", "vars", "varsaw", "querystorage", "tokens", "logins", "bingads", "varsac", "modules", "proxysettings", "cred", "keywordtool", "codes", "facebook", "qt", "youtube", "settings", "flickr", "twitter", "webmaster", "stripe", "fbads", "fbinsights", "mailchimp"
        isSheetAconfigSheet = True
    Case Else
        isSheetAconfigSheet = False
    End Select

End Function


Sub runAllQueriesInQueryStorage()
    Dim vsar As Long
    Dim col As Long
    With Sheets("querystorage")
        vsar = vikasar(.Cells(Range("queryTypeRow").row, 1))
        For col = 3 To vsar
            Call runQueryFromQueryStorageCol(col)
        Next col
    End With
End Sub

Sub runSelectedQueriesInQueryStorage()
    Dim col As Long
    For col = Selection.Column To Selection.Column + Selection.Columns.Count - 1
        Call runQueryFromQueryStorageCol(col)
    Next col
End Sub




Sub runQueryFromQueryStorageCol(col As Long)

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    runningSheetRefresh = True
    Dim sheetRow As Long
    Dim vsar As Long

    Dim sheetID As String
    Dim sheetName As String
    Dim sheetIDRow As Long
    Dim useSheetDates As Boolean

    Dim sDate As Variant
    Dim eDate As Variant

    Dim sDateSaved As Variant
    Dim eDateSaved As Variant

    sheetIDRow = Range("querySheetIDRow").row

    sheetRow = Range("querySheetRow").row

    With Sheets("querystorage")

        sheetID = .Cells(sheetIDRow, col).value

        'if no ID found then create one
        If sheetID = vbNullString Then
            sheetID = createNewSheetID()
            .Cells(sheetIDRow, col).value = sheetID
            .Cells(sheetIDRow, col).Name = sheetID
        End If


        sheetName = findSheetNameForSheetID(sheetID)
        If sheetName = vbNullString Then sheetName = .Cells(sheetRow, col).value

        .Cells(sheetRow, col).value = sheetName

        If sheetName <> vbNullString And SheetExists(sheetName) Then
            Sheets(sheetName).Select
            Call refreshDataOnSelectedSheet
        Else

            dataSource = .Cells(Range("datasourceRow").row, col).value

            'saves current inteface query
            Call markToCurrentQuery
            Call saveQueryFromCQ


            If sheetName = vbNullString Then
                'if no sheet found then create a name
                Dim sheetNum As Long
                For sheetNum = 1 To 1000
                    If Not SheetExists("report" & sheetNum) Then
                        .Cells(sheetRow, col).value = "report" & sheetNum
                        Exit For
                    End If
                Next sheetNum
            End If

            If importingFromOldVersion = False And SheetExists(sheetName) = False Then runningSheetRefresh = False

            .Cells(Range("runDateRow").row, col).value = Now()

            dataSource = .Cells(Range("datasourceRow").row, col).value

            Call setDatasourceVariables

            dateRangeType = .Cells(Range("dateRangeTypeRow").row, col).value

            If dateRangeType = "custom" Or dateRangeType = vbNullString Then dateRangeType = "fixed"

            If dateRangeType = "fixed" Or dateRangeType = "custom" Then
                With Sheets("querystorage").Cells(Range("sdateRowQS").row, col)

                    useSheetDates = True
                    If .value <> "" And .Offset(1).value <> "" Then

                        On Error GoTo qsdateError
                        useSheetDates = False
                        If IsDate(CDate(.value)) Then sDate = CDate(.value)
                        If IsDate(CDate(.Offset(1).value)) Then eDate = CDate(.Offset(1).value)
                        On Error Resume Next
                        If debugMode = True Then On Error GoTo 0

                    End If

                    If useSheetDates = True Then
                        sDate = Range("startDate" & varsuffix).value
                        eDate = Range("endDate" & varsuffix).value
                        On Error Resume Next
                        If debugMode = True Then On Error GoTo 0
                    End If

                End With
            Else
                Call getDatesForDateRangeType(dateRangeType)
                sDate = startDate
                eDate = endDate
            End If

            sDateSaved = Range("startdate" & varsuffix).value
            eDateSaved = Range("enddate" & varsuffix).value


            Range("startDate" & varsuffix).value = sDate
            Range("endDate" & varsuffix).value = eDate

            .Columns(ColumnLetter(col)).Copy Sheets("vars").Columns(ColumnLetter(Range("parameterListStart").Column + 1))
            Call getFromCurrentQuery(, True)

            Application.DisplayAlerts = False
            If SheetExists(.Cells(sheetRow, col).value) = True And .Cells(Range("deleteSheetRow").row, col).value = True Then Sheets(.Cells(sheetRow, col).value).Delete

            Select Case Range("queryType").value
            Case "A"
                Call fetchAggregateFigures
            Case Else
                Call fetchFiguresSplitByDimensions
            End Select


            Range("startdate" & varsuffix).value = sDateSaved
            Range("enddate" & varsuffix).value = eDateSaved

            Call returnSavedQueryToCQ

        End If

    End With

    Exit Sub

qsdateError:
    useSheetDates = True
End Sub


Sub refreshDataOnAllSheetsDontOverrideDates()

    On Error Resume Next
    Dim sh As Worksheet
    Dim foundReportToRefresh As Boolean

    Dim sheetWasHidden As Boolean
    Dim callerSheet As Worksheet

    Set callerSheet = ActiveSheet

    foundReportToRefresh = False

    Dim sheetID As String


    For Each sh In ThisWorkbook.Worksheets

        If isSheetAconfigSheet(sh.Name) = False Then
            If sh.Visible = False Then
                sheetWasHidden = True
                sh.Visible = True
                Application.ScreenUpdating = False
            Else
                sheetWasHidden = False
                Application.ScreenUpdating = True
            End If

            sh.Select


            sheetID = sh.Cells(1, 1).value
            sheetID = findRangeName(sh.Cells(1, 1))

            If Left(sheetID, 3) <> "_SH" Then sheetID = ""

            If sheetID = vbNullString Then
                sheetID = findSheetIDForSheetName(sh.Name)
            End If

            If sheetID <> vbNullString And queryExistsInQueryStorage(sheetID) = True Then
                sh.Cells(1, 1).value = sheetID
                Call refreshDataOnSelectedSheet
                foundReportToRefresh = True
                Application.ScreenUpdating = True
            ElseIf sheetID <> vbNullString Then
                MsgBox "The data for this sheet (" & sh.Name & ") could not be refreshed, as the query could not be found from the querystorage sheet. You need to run the query again through the interface in the config sheet."
            End If

            If sheetWasHidden = True Then
                sh.Visible = False
            End If
        End If

    Next sh

    If Range("updatePivotTables").value Then Call refreshPivotTables

    callerSheet.Select

    If foundReportToRefresh = True Then
        '   If creatingClientFiles = False Then MsgBox "Refreshing reports done!"
    Else
        MsgBox "Found no reports to refresh."
    End If

End Sub



Sub selectActiveReportInQuerystorage()


    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim col As Long

    Dim thisSheet As Worksheet
    Set thisSheet = ActiveSheet


    Dim sheetID As String
    sheetID = thisSheet.Cells(1, 1).value
    sheetID = findRangeName(thisSheet.Cells(1, 1))

    If sheetID = vbNullString Then
        sheetID = findSheetIDForSheetName(ActiveSheet.Name)
        thisSheet.Cells(1, 1).value = sheetID
        If sheetID <> vbNullString Then thisSheet.Cells(1, 1).Name = sheetID
    End If


    If sheetID = vbNullString Or queryExistsInQueryStorage(sheetID) = False Then
        MsgBox "This report cannot be modified, as the query could not be found from the querystorage sheet. You need to run the query again through the query builder interface."
        Exit Sub
    End If

    On Error GoTo errhandler
    col = Application.Match(sheetID, Sheets("querystorage").Rows(Range("querySheetIDrow").row), 0)
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0


    With Sheets("querystorage")
        .Visible = xlSheetVisible
        .Select
        .Cells(1, col).EntireColumn.Select
    End With

    Exit Sub

errhandler:

    MsgBox "This report cannot be modified, as the query could not be found from the querystorage sheet. You need to run the query again through the query builder interface."
    Application.EnableEvents = True
    Exit Sub

End Sub





Sub markToCurrentQuery()

'marks query selected through interface into currentquery
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim varsSheet As Worksheet
    Set varsSheet = Sheets("vars")


    Call setDatasourceVariables


    Dim rivi As Long
    Dim col As Long

    Dim dimensionsListStartRow As Long
    Dim dimensionsListStartColumn As Long

    Dim metricsListStartRow As Long
    Dim metricsListStartColumn As Long


    With Range("dimensionsAllStart")
        dimensionsListStartRow = .row
        dimensionsListStartColumn = .Column
    End With

    With Range("metricsAllStart")
        metricsListStartRow = .row
        metricsListStartColumn = .Column
    End With


    Dim dimensionsNumsStartRow As Long
    Dim dimensionsNumsStartcolumn As Long

    Dim metricsNumsStartRow As Long
    Dim metricsNumsStartcolumn As Long

    With Range("dnum1")
        dimensionsNumsStartRow = .row
        dimensionsNumsStartcolumn = .Column
    End With

    With Range("mnum1")
        metricsNumsStartRow = .row
        metricsNumsStartcolumn = .Column
    End With



    Dim profileListcolumn As Long


    Dim arvo As Variant
    Dim arvo2 As Variant

    Dim listarivi As Long

    Dim profRiviConfig As Long
    Dim profRiviVars As Long

    Dim valuesArr As Variant
    Dim valuesArrCQcol As Variant
    Dim valuesRng As Range

    Dim valuesArrProf As Variant
    Dim valuesRngProf As Range


    col = Range("parameterListStart").Column

    With Sheets("vars")


        queryType = Range("queryType").value
        .Cells(1, col + 1).EntireColumn.ClearContents
        Range("queryType").value = queryType

        If Application.Calculation <> xlAutomatic Then
            varsSheetForDataSource.Calculate
            '.Cells.Dirty
            .Calculate
        End If

        If dataSource = "GA" Then
            Range("filterstring").Offset(, 1).value = Range("filterstringGA").value
        Else
            Range("filterstring").Offset(, parameterColumnOffset - 1).value = Range("filterstring" & varsuffix).value
        End If



        If dataSource <> "TW" Then

            Set profileListStart = Range("profileListStart" & varsuffix)
            profileListcolumn = profileListStart.Column

            With configsheet
                Set valuesRngProf = profileListStart.Offset(, -2).Resize(Range("profiles" & varsuffix).Rows.Count, 5)
            End With
            valuesArrProf = valuesRngProf.value

            Set valuesRng = .Range(.Cells(Range("parameterListStart").row, col), .Cells(vikarivi(Range("parameterListStart")) + Application.CountA(valuesRngProf.Cells(1, 1).EntireColumn), col + parameterColumnOffset))
        Else
            Set valuesRng = .Range(.Cells(Range("parameterListStart").row, col), .Cells(vikarivi(Range("parameterListStart")), col + parameterColumnOffset))

        End If
        valuesArr = valuesRng.value

        ReDim valuesArrCQcol(LBound(valuesArr) To UBound(valuesArr), 1 To 1)

        col = 1


        For rivi = LBound(valuesArr) To UBound(valuesArr)

            arvo = valuesArr(rivi, col)
            arvo2 = valuesArr(rivi, col + parameterColumnOffset)

            If arvo = "Dimensions" Then
                If dataSource = "TW" Then Exit For
                For listarivi = rivi To rivi + 11

                    arvo2 = valuesArr(listarivi, col + parameterColumnOffset)

                    If fieldNameIsOk(arvo2) = True Then valuesArr(listarivi, col + 1) = arvo2

                Next listarivi

            ElseIf arvo = "Metrics" Then

                For listarivi = rivi To rivi + 11

                    arvo2 = valuesArr(listarivi, col + parameterColumnOffset)

                    If fieldNameIsOk(arvo2) = True Then valuesArr(listarivi, col + 1) = arvo2

                Next listarivi

            ElseIf arvo = "Profiles" Then

                profRiviVars = rivi

                For profRiviConfig = LBound(valuesArrProf) To UBound(valuesArrProf)
                    If valuesArrProf(profRiviConfig, 1) <> vbNullString Then
                        valuesArr(profRiviVars, col + 1) = valuesArrProf(profRiviConfig, 5)
                        profRiviVars = profRiviVars + 1
                    End If
                Next profRiviConfig

                Erase valuesArrProf

                Exit For

            ElseIf arvo = "Date range type" Then

                If InStr(1, arvo2, "lastx") > 0 Then
                    If IsNumeric(configsheet.Shapes("lastXbox" & varsuffix).TextFrame.Characters.Text) Then
                        If (configsheet.Shapes("lastXbox" & varsuffix).TextFrame.Characters.Text > 0) Then
                            arvo2 = Replace(arvo2, "lastx", "last" & val(configsheet.Shapes("lastXbox" & varsuffix).TextFrame.Characters.Text))
                            If configsheet.CheckBoxes("includeCurrentCB").value = 1 Then arvo2 = arvo2 & "inc"
                        End If
                    End If
                End If

                valuesArr(rivi, col + 1) = arvo2
            ElseIf dataSource = "TW" And arvo = "Columns" Then
                arvo2 = ""
                With Twitter
                    If .CheckBoxes("timeCB").value = 1 Then
                        arvo2 = arvo2 & "time,"
                    End If
                    If .CheckBoxes("twitter_nameCB").value = 1 Then
                        arvo2 = arvo2 & "twitter_name,"
                    End If
                    If .CheckBoxes("nameCB").value = 1 Then
                        arvo2 = arvo2 & "name,"
                    End If
                    If .CheckBoxes("locationCB").value = 1 Then
                        arvo2 = arvo2 & "location,"
                    End If
                    If .CheckBoxes("followersCB").value = 1 Then
                        arvo2 = arvo2 & "followers,"
                    End If
                    If .CheckBoxes("tweetCB").value = 1 Then
                        arvo2 = arvo2 & "tweet,"
                    End If
                    If .CheckBoxes("languageCB").value = 1 Then
                        arvo2 = arvo2 & "language,"
                    End If
                    If .CheckBoxes("retweetsCB").value = 1 Then
                        arvo2 = arvo2 & "retweets,"
                    End If
                    If .CheckBoxes("linkCB").value = 1 Then
                        arvo2 = arvo2 & "link,"
                    End If
                End With
                valuesArr(rivi, col + 1) = arvo2
            ElseIf arvo <> vbNullString And arvo <> "Query type" Then

                valuesArr(rivi, col + 1) = arvo2

            End If



        Next rivi


        For rivi = LBound(valuesArr) To UBound(valuesArr)
            valuesArrCQcol(rivi, 1) = valuesArr(rivi, 2)
        Next rivi

        valuesRng.Columns("B").value = valuesArrCQcol
        Erase valuesArr

        Range("dataSource").value = dataSource


        Dim sheetID As String
        sheetID = createNewSheetID()
        Range("sheetID").value = sheetID

        Dim sheetNum As Long
        Dim foundFreeSheetNum As Boolean

        For sheetNum = 1 To 1000
            If Not SheetExists("report" & sheetNum) Then
                Range("wsName").value = "report" & sheetNum
                foundFreeSheetNum = True
                Exit For
            End If
        Next sheetNum
        If Not foundFreeSheetNum Then
            sheetNum = Round(Rnd * 1000000000, 0)
            Range("wsName").value = "report0" & sheetNum
        End If


    End With


End Sub




Sub getFromCurrentQuery(Optional dontCopyProfiles As Boolean = False, Optional dontUpdateDropdowns As Boolean = False)

'gets currentquery ready for execution
    Dim aika
    aika = Timer

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Call unprotectSheets


    dataSource = Range("dataSource").value

    Call setDatasourceVariables

    If dataSource = "TW" Then
        Call protectSheets
        Exit Sub
    End If


    Dim rivi As Long
    Dim rivi2 As Long
    Dim listarivi As Long
    Dim col As Long

    Dim dimensionsListStartRow As Long
    Dim dimensionsListStartColumn As Long

    Dim metricsListStartRow As Long
    Dim metricsListStartColumn As Long

    dimensionsListStartRow = Range("dimensionsAllStart" & varsuffix).row
    dimensionsListStartColumn = Range("dimensionsAllStart" & varsuffix).Column

    metricsListStartRow = Range("metricsAllStart" & varsuffix).row
    metricsListStartColumn = Range("metricsAllStart" & varsuffix).Column


    Dim dimensionsNumsStartRow As Long
    Dim dimensionsNumsStartcolumn As Long

    Dim metricsNumsStartRow As Long
    Dim metricsNumsStartcolumn As Long

    dimensionsNumsStartRow = Range("dnum1" & varsuffix).row
    dimensionsNumsStartcolumn = Range("dnum1" & varsuffix).Column

    metricsNumsStartRow = Range("mnum1" & varsuffix).row
    metricsNumsStartcolumn = Range("mnum1" & varsuffix).Column


    Set profileListStart = Range("profileListStart" & varsuffix)

    Dim profileListcolumn As Long
    profileListcolumn = profileListStart.Column

    Dim arvo As Variant
    Dim arvo2 As Variant

    Dim profRiviConfig As Long
    Dim profRiviVars As Long

    Dim valuesArr As Variant
    Dim valuesRng As Range

    Dim valuesArrTemp As Variant
    Dim resultValues As Variant

    col = Range("parameterListStart").Column

    Dim profNumsConfig As Range
    '  Set profNumsConfig = configsheet.Range(ColumnLetter(profileListcolumn + 2) & profileListStart.Row & ":" & ColumnLetter(profileListcolumn + 2) & vikarivi(profileListStart.Offset(, 2)))
    With configsheet
        Set profNumsConfig = .Range(.Cells(profileListStart.row, profileListcolumn + 2), .Cells(vikarivi(profileListStart.Offset(, 2)), profileListcolumn + 2))
    End With

    Dim vriviProfs As Long
    vriviProfs = vikarivi(profileListStart.Offset(, 2))
    With Sheets("vars")


        Set valuesRng = .Range(.Cells(Range("parameterListStart").row, col), .Cells(vikarivi(Range("parameterListStart").Offset(, 1)), col + 1))
        valuesArr = valuesRng.value
        col = 1

        For rivi = LBound(valuesArr) To UBound(valuesArr)

            arvo = valuesArr(rivi, col)

            If arvo = "Dimensions" Then

                valuesArrTemp = Range("dimensionsAllStart" & varsuffix).Resize(vikarivi(Range("dimensionsAllStart" & varsuffix))).value
                ReDim resultValues(1 To 12, 1 To 1)

                For listarivi = rivi To rivi + 11

                    arvo2 = valuesArr(listarivi, col + 1)

                    If fieldNameIsOk(arvo2) = True Then

                        rivi2 = 1
                        On Error Resume Next
                        rivi2 = Application.Match(arvo2, valuesArrTemp, 0)

                        resultValues(listarivi - rivi + 1, 1) = rivi2

                    Else
                        resultValues(listarivi - rivi + 1, 1) = 1

                    End If

                Next listarivi

                Range("dnum1" & varsuffix).Resize(UBound(resultValues)).value = resultValues
                Erase valuesArrTemp
                Erase resultValues

            ElseIf arvo = "Metrics" Then

                valuesArrTemp = Range("metricsAllStart" & varsuffix).Resize(vikarivi(Range("metricsAllStart" & varsuffix))).value
                ReDim resultValues(1 To 12, 1 To 1)

                For listarivi = rivi To rivi + 11

                    arvo2 = valuesArr(listarivi, col + 1)

                    If fieldNameIsOk(arvo2) = True Then

                        rivi2 = 1
                        On Error Resume Next
                        rivi2 = Application.Match(arvo2, valuesArrTemp, 0)
                        resultValues(listarivi - rivi + 1, 1) = rivi2
                        '      varsSheetForDataSource.Cells(metricsNumsStartRow + listarivi - rivi, metricsNumsStartcolumn).value = rivi2

                    Else
                        resultValues(listarivi - rivi + 1, 1) = 1
                        '  varsSheetForDataSource.Cells(metricsNumsStartRow + listarivi - rivi, metricsNumsStartcolumn).value = 1

                    End If
                Next listarivi

                Range("mnum1" & varsuffix).Resize(UBound(resultValues)).value = resultValues
                Erase valuesArrTemp
                Erase resultValues

            ElseIf arvo = "Segmenting dimension" Or arvo = "Segmenting dimension 2" Then

                arvo2 = valuesArr(rivi, col + 1)

                If fieldNameIsOk(arvo2) = True Then

                    rivi2 = 1
                    On Error Resume Next

                    rivi2 = Application.Match(arvo2, Range("sdimensionsAllStart" & varsuffix).Resize(1000), 0)
                    If arvo = "Segmenting dimension" Then
                        Range("sdnum1").value = rivi2
                    Else
                        Range("sdnum2").value = rivi2
                    End If

                Else
                    If arvo = "Segmenting dimension" Then
                        Range("sdnum1").value = 1
                    Else
                        Range("sdnum2").value = 1
                    End If

                End If

            ElseIf arvo = "Profiles" And dontCopyProfiles = False Then



                Exit For


            End If

        Next rivi

        If Application.Calculation <> xlAutomatic Then
            varsSheetForDataSource.Calculate
            If varsSheetForDataSource <> vars Then
                .Calculate
            End If
        End If

    End With

    If Not dontUpdateDropdowns Then
        If dataSource = "GA" Then
            Call updateFieldSelectionsGA
        ElseIf dataSource = "AW" Then
            Call updateFieldSelectionsAW
        ElseIf dataSource = "AC" Then
            Call updateFieldSelectionsAC
        ElseIf dataSource = "FB" Then
            Call updateFieldSelectionsFB
        ElseIf dataSource = "YT" Then
            Call updateFieldSelectionsYT
        End If
    End If

    Call protectSheets

    Debug.Print "AIKA:" & Timer - aika

End Sub




Sub copyCurrentquerytoQueryStorage()

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Dim queryStoragecol As Long
    queryStoragecol = 1 + vikasar(Sheets("querystorage").Cells(Range("queryTypeRow").row, 1))
    If queryStoragecol > columnLimit Then Exit Sub

    Sheets("vars").Columns(ColumnLetter(Range("parameterListStart").Column + 1)).Copy Sheets("querystorage").Columns(ColumnLetter(queryStoragecol))




End Sub



Public Function queryExistsInQueryStorage(sheetID As String) As Boolean
    If Application.CountIf(Sheets("querystorage").Rows(Range("querysheetidrow").row), sheetID) > 0 Then
        queryExistsInQueryStorage = True
    Else
        queryExistsInQueryStorage = False
    End If
End Function




Sub deleteFromQueryStorage(querySheetName As String)

    Application.ScreenUpdating = False

    On Error Resume Next

    Dim col As Long
    col = 0
    With Sheets("queryStorage")

        col = Application.Match(querySheetName, .Rows(Range("querysheetrow").row), 0)
        If col <> 0 Then .Columns(ColumnLetter(col)).Delete

    End With
    If debugMode = True Then On Error GoTo 0
End Sub



Sub deleteFromQueryStorageByID(querySheetID As String)

    Application.ScreenUpdating = False

    On Error Resume Next

    Dim col As Long
    col = 0
    With Sheets("queryStorage")

        col = Application.Match(querySheetID, .Rows(Range("querysheetidrow").row), 0)
        If col <> 0 Then .Columns(ColumnLetter(col)).Delete

    End With
    If debugMode = True Then On Error GoTo 0
End Sub


Public Function findSheetNameForSheetID(sheetID As String, Optional wb As Workbook) As String

    On Error Resume Next
    'On Error GoTo 0
    findSheetNameForSheetID = vbNullString
    Dim sh As Worksheet
    Dim foundMatch As Boolean

    If IsMissing(wb) = True Or wb Is Nothing Then Set wb = ThisWorkbook

    For Each sh In wb.Worksheets

        foundMatch = False

        If findRangeName(sh.Cells(1, 1), wb) = sheetID Then foundMatch = True

        If sh.Cells(1, 1).value = sheetID Then foundMatch = True

        If foundMatch = True Then
            findSheetNameForSheetID = sh.Name
            Exit Function
        End If

    Next sh

    findSheetNameForSheetID = vbNullString

End Function



Public Function findSheetIDForSheetName(sheetName As String, Optional wb As Workbook) As String


    On Error Resume Next
    findSheetIDForSheetName = vbNullString
    Dim col As Long
    col = 0

    If IsMissing(wb) = True Or wb Is Nothing Then Set wb = ThisWorkbook

    With wb.Sheets("queryStorage")

        col = Application.Match(sheetName, .Rows(Range("querysheetrow").row), 0)

        If col <> 0 Then findSheetIDForSheetName = .Cells(Range("querysheetidrow").row, col).value

    End With

    If debugMode = True Then On Error GoTo 0

End Function

Sub updateSheetNamesInQueryStorage()


    Dim col As Long
    Dim sheetRow As Long
    Dim vsar As Long

    Dim sheetID As String
    Dim sheetIDRow As Long
    sheetIDRow = Range("querySheetIDRow").row
    sheetID = 0

    sheetRow = Range("querySheetRow").row

    With Sheets("querystorage")

        vsar = vikasar(.Cells(Range("queryTypeRow").row, 1))

        For col = 3 To vsar

            sheetID = .Cells(sheetIDRow, col).value

            If sheetID <> vbNullString Then

                .Cells(sheetRow, col).value = findSheetNameForSheetID(sheetID)

            End If

        Next col

    End With

End Sub


