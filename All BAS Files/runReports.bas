Attribute VB_Name = "runReports"
Option Private Module
Option Explicit

Sub runReport()
    On Error Resume Next
    If debugMode Then On Error GoTo 0
    Application.ScreenUpdating = False

    Dim profileArrLoc As Variant
    Dim i As Long
    Dim dimensionNum As Long
    Dim rivi As Long
    Dim profileCountLoc As Long
    Dim profileNumLoc As Long
    Dim profIDloc As String
    Dim selectedProfileArrLoc As Variant
    Dim dimensionsListStart As Range
    Dim metricsListStart As Range
    Dim dimensionsFound As Boolean
    Dim metricsFound As Boolean
    Dim sdQuery As Boolean
    Dim separateReportForEachProfile As Boolean
    Dim profileStr As String

    calculationSetting = Application.Calculation

    Call checkOperatingSystem

    If usingMacOSX = False Then ProgressBox.Show False
    Call updateProgress(1, "Starting report run...", , False)

    Call setDatasourceVariables
    '  Call testConnection

    If usingMacOSX Then
        Analytics.Shapes("fieldsOKnote").Visible = False
        AdWords.Shapes("fieldsOKnote").Visible = False
        BingAds.Shapes("fieldsOKnote").Visible = False
    End If

    profileSelectionsArr = Range("profileSelections" & varsuffix).value

    If Not IsArray(profileSelectionsArr) Then
        profileStr = profileSelectionsArr
        ReDim profileSelectionsArr(1 To 1, 1 To 1)
        profileSelectionsArr(1, 1) = profileStr
        profileStr = vbNullString
    End If

    profileCountLoc = 0
    For rivi = 1 To UBound(profileSelectionsArr)
        If profileSelectionsArr(rivi, 1) <> vbNullString Then
            profileCountLoc = profileCountLoc + 1
        End If
    Next rivi

    If profileCountLoc = 0 Then
        Application.StatusBar = False
        MsgBox "No " & referToProfilesAs & " have been selected. Select at least one from the list and try again."
        Call hideProgressBox
        End
    End If

    If Application.Calculation <> xlAutomatic Then
        varsSheetForDataSource.Calculate
        vars.Calculate
    End If

    profileArrLoc = Range("profiles" & varsuffix).value

    profileNumLoc = 0
    ReDim selectedProfileArrLoc(1 To profileCountLoc, 1 To 2)
    For rivi = 1 To UBound(profileSelectionsArr)
        If profileSelectionsArr(rivi, 1) <> vbNullString Then
            profileNumLoc = profileNumLoc + 1
            selectedProfileArrLoc(profileNumLoc, 1) = profileArrLoc(rivi, 5)
            selectedProfileArrLoc(profileNumLoc, 2) = rivi
        End If
    Next rivi

    DoEvents

    Set metricsListStart = Range("metric1name" & varsuffix)
    metricsFound = False
    With metricsListStart.Worksheet
        For metricNum = 1 To 12
            If fieldNameIsOk(.Cells(metricNum + metricsListStart.row - 1, metricsListStart.Column).value) = True Then
                metricsFound = True
                Exit For
            End If
        Next metricNum
    End With

    If metricsFound = False And dataSource <> "MC" And dataSource <> "FA" Then
        Application.StatusBar = False
        MsgBox "Choose at least one metric first"
        Call hideProgressBox
        End
    End If

    Call updateProgress(2, "Starting report run...", , False)
    DoEvents

    Set dimensionsListStart = Range("dimension1name" & varsuffix)
    dimensionsFound = False
    With dimensionsListStart.Worksheet
        For dimensionNum = 1 To 10
            If fieldNameIsOk(.Cells(dimensionNum + dimensionsListStart.row - 1, dimensionsListStart.Column).value) = True Then
                dimensionsFound = True
                Exit For
            End If
        Next dimensionNum
    End With



    sdQuery = False
    If Application.Calculation <> xlAutomatic Then
        Range("segmDimension1name" & varsuffix).Calculate
        Range("segmDimension2name" & varsuffix).Calculate
    End If
    If fieldNameIsOk(Range("segmDimension1name" & varsuffix).value) Or fieldNameIsOk(Range("segmDimension2name" & varsuffix).value) Then
        dimensionsFound = True
        sdQuery = True
    End If

    DoEvents



    runningSheetRefresh = False
    importingFromOldVersion = False

    If dataSource = "FB" Then
        reportRunType = "combined"
    ElseIf profileCountLoc > 1 Then
        reportRunType = "combined"
        Call hideProgressBox
        With reportTypeUF
            .Label1.Caption = "You've selected " & profileCountLoc & " " & referToProfilesAs & ". What kind of reporting do you prefer?"
            .combinedB.Caption = "COMBINED" & vbCrLf & "One report, each " & referToProfilesAsSing & " displayed separately"
            .summedB.Caption = "SUMMED" & vbCrLf & "One report, all " & referToProfilesAs & " summed up"
            .separateB.Caption = "SEPARATE" & vbCrLf & "Separate report for each " & referToProfilesAsSing
            .Show
        End With
    Else
        reportRunType = "combined"
    End If

    Call markToCurrentQuery

    If Range("reportFormattingType").value = 1 Then
        rawDataReport = True
    Else
        rawDataReport = False
    End If

    Call updateProgress(3, "Starting report run...", , False)

    If dimensionsFound = False And Not rawDataReport Then
        runningMultipleReports = False
        DoEvents
        Range("queryType").value = "A"
        If reportRunType = "summed" Then
            Range("sumAllProfiles").value = True
        Else
            Range("sumAllProfiles").value = False
        End If
        Range("queryRunTime").value = Now()
        Call copyCurrentquerytoQueryStorage
        DoEvents
        Call fetchAggregateFigures
    Else
        If reportRunType = "combined" Or reportRunType = "summed" Then
            runningMultipleReports = False
            If sdQuery = False Then
                Range("queryType").value = "D"
            Else
                Range("queryType").value = "SD"
            End If
            If reportRunType = "summed" Then
                Range("sumAllProfiles").value = True
            Else
                Range("sumAllProfiles").value = False
            End If
            Range("queryRunTime").value = Now()
            DoEvents
            Call copyCurrentquerytoQueryStorage
            DoEvents
            Call fetchFiguresSplitByDimensions
        Else
            runningMultipleReports = True
            Range("sumAllProfiles").value = False
            For profileNumLoc = 1 To profileCountLoc
                Call markToCurrentQuery
                If sdQuery = False Then
                    Range("queryType").value = "D"
                Else
                    Range("queryType").value = "SD"
                End If

                profIDloc = selectedProfileArrLoc(profileNumLoc, 1)
                With Range("profileSelections" & varsuffix)
                    .ClearContents
                    .Cells(selectedProfileArrLoc(profileNumLoc, 2), 1).value = profIDloc
                End With
                Range("profilesStartCQ").Resize(1000, 1).ClearContents
                Range("profilesStartCQ") = profIDloc
                Range("queryRunTime").value = Now()
                DoEvents
                Call copyCurrentquerytoQueryStorage
                DoEvents
                Call fetchFiguresSplitByDimensions
            Next profileNumLoc
            Range("profileSelections" & varsuffix).value = profileSelectionsArr
        End If
    End If

End Sub

Sub showQuestionUF(question As String, b1 As String, b2 As String, header As String, Optional textInput As Boolean = False)
    With questionUF
        .Caption = header
        .Label1.Caption = question
        .b1.Caption = b1
        .b2.Caption = b2
        If textInput = True Then
            .b2.Visible = False
            .TextBox1.Visible = True
        End If
        .Show
    End With
End Sub

Sub aggregateQuery()
    stParam1 = "8"
    stParam2 = "A"

    runningSheetRefresh = False
    importingFromOldVersion = False
    Application.ScreenUpdating = False
    Call markToCurrentQuery
    Range("queryType").value = "A"
    Range("queryRunTime").value = Now()
    Call copyCurrentquerytoQueryStorage
    Call fetchAggregateFigures
End Sub

Sub dimensionQuery()
    stParam1 = "8"
    stParam2 = "D"
    runningSheetRefresh = False
    importingFromOldVersion = False
    Application.ScreenUpdating = False

    Call markToCurrentQuery

    '-----------------------------------------
    'determine whether query type is D or SD
    segmDimCount = 0
    If fieldNameIsOk(Range("segmDimName").value) = True Then segmDimCount = segmDimCount + 1
    If fieldNameIsOk(Range("segmDimName2").value) = True Then segmDimCount = segmDimCount + 1

    If segmDimCount = 0 Then
        Range("queryType").value = "D"
    Else
        Range("queryType").value = "SD"
    End If
    '-----------------------------------------
    Range("queryRunTime").value = Now()

    Call copyCurrentquerytoQueryStorage
    Call fetchFiguresSplitByDimensions

End Sub


