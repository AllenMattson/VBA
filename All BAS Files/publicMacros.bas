Attribute VB_Name = "publicMacros"
Option Explicit

Sub refreshDataOnAllSheets()

    Call refreshDataOnAllSheetsDontOverrideDates

End Sub

Sub refreshPivotTables()
    On Error Resume Next
    Dim w As Worksheet, p As PivotTable
    For Each w In ThisWorkbook.Worksheets
        For Each p In w.PivotTables
            p.RefreshTable
            p.Update
        Next
    Next
End Sub


Sub refreshDataOnAllSheetsAndPivotTables()
    Call refreshDataOnAllSheetsDontOverrideDates
    Call refreshPivotTables
End Sub



Sub refreshDataOnSelectedSheet()


    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    calculationSetting = Application.Calculation
    '   Call testConnection



    runningSheetRefresh = True
    importingFromOldVersion = False
    Application.EnableEvents = False

    Application.ScreenUpdating = False

    Call checkOperatingSystem

    stParam1 = "9"
    stParam2 = CStr(usingMacOSX)

    If usingMacOSX = False Then ProgressBox.Show False
    Call updateProgress(1, "Starting report refresh...", , False)


    If isSheetAconfigSheet(ActiveSheet.Name) = True Then Exit Sub
    Dim col As Long

    Dim sDate As Variant
    Dim eDate As Variant

    Dim sDateSaved As Variant
    Dim eDateSaved As Variant

    Dim useSheetDates As Boolean

    Dim thisSheet As Object
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
        reportRunSuccessful = False
        MsgBox "The data for this sheet could not be refreshed, as the query could not be found from the querystorage sheet. You need to run the query again through the query builder interface."
        Call hideProgressBox
        Exit Sub
    End If


    On Error GoTo errhandler
    col = Application.Match(sheetID, Sheets("querystorage").Rows(Range("querySheetIDrow").row), 0)
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    dataSource = Sheets("querystorage").Cells(Range("dataSourceRow").row, col).value
    If dataSource = "" Then dataSource = "GA"

    Call setDatasourceVariables


    If dataSource <> "TW" Then
        dateRangeType = Sheets("querystorage").Cells(Range("dateRangeTypeRow").row, col).value

        If dateRangeType = "custom" Or dateRangeType = vbNullString Then dateRangeType = "fixed"

        If dateRangeType = "fixed" Then
            With Sheets("querystorage").Cells(Range("sdateRowQS").row, col)

                useSheetDates = True
                If .value <> "" And .Offset(1).value <> "" Then
                    ' If IsDate(CDate(.value)) And IsDate(CDate(.Offset(1).value)) Then
                    On Error GoTo qsdateError
                    useSheetDates = False
                    If IsDate(CDate(.value)) Then sDate = CDate(.value)
                    If IsDate(CDate(.Offset(1).value)) Then eDate = CDate(.Offset(1).value)

                    On Error Resume Next
                    If debugMode = True Then On Error GoTo 0
                    '   End If
                End If

                If useSheetDates = True Then
                    On Error GoTo dateError
                    sDate = Range(sheetID & "_sdate").value
                    eDate = Range(sheetID & "_edate").value
                    On Error Resume Next
                    If debugMode = True Then On Error GoTo 0
                End If

            End With
        Else
            Call getDatesForDateRangeType(dateRangeType)
            sDate = startDate
            eDate = endDate
        End If


        If sDate > eDate Then
            MsgBox "Invalid date range (start date should be before end date)"
            Call hideProgressBox
            Exit Sub
        End If

    End If




    If Range("loggedin" & varsuffix).value = False And configsheet.Visible <> xlSheetVisible Then
        MsgBox "You need to be logged in to run reports. Log in and try again."
        Call hideProgressBox
        End
    End If


    'saves current inteface query
    Call markToCurrentQuery
    Call saveQueryFromCQ

    Sheets("querystorage").Cells(Range("querySheetRow").row, col).value = findSheetNameForSheetID(sheetID)
    '   tempArr = Sheets("querystorage").Cells(1, col).Resize(20000, 1).value
    Range("parameterListStart").Offset(, 1).EntireColumn.Cells(1, 1).Resize(20000, 1).value = Sheets("querystorage").Cells(1, col).Resize(20000, 1).value
    '  Sheets("querystorage").Columns(ColumnLetter(col)).Copy Sheets("vars").Columns(ColumnLetter(Range("parameterListStart").Column + 1))
    'Call copyValues(Sheets("querystorage").Cells(1, col).Resize(5000), Sheets("vars").Cells(1, Range("parameterListStart").Column + 1))

    Call getFromCurrentQuery(, True)

    Call setDatasourceVariables


    If dataSource <> "TW" Then
        sDateSaved = Range("startdate" & varsuffix).value
        eDateSaved = Range("enddate" & varsuffix).value

        Range("startdate" & varsuffix).value = sDate
        Range("enddate" & varsuffix).value = eDate


        If Range("deleteSheetOnRefresh").value = True Then
            Application.DisplayAlerts = False
            thisSheet.Delete
        End If




        Select Case Range("queryType").value
        Case "A"
            Call fetchAggregateFigures
        Case Else
            Call fetchFiguresSplitByDimensions
        End Select
        Range("startdate" & varsuffix).value = sDateSaved
        Range("enddate" & varsuffix).value = eDateSaved
    Else
        Call fetchTweets
    End If



    Call returnSavedQueryToCQ


    Application.EnableEvents = True
    Call getFromCurrentQuery



    Exit Sub

errhandler:

    MsgBox "The data for this sheet could not be refreshed, as the query could not be found from the querystorage sheet. You need to run the query again through the query builder interface."
    Application.EnableEvents = True
    Exit Sub

dateError:
    MsgBox "The date range is invalid, query can not be refreshed."
    Application.EnableEvents = True
    Exit Sub

qsdateError:
    useSheetDates = True

End Sub

