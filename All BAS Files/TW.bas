Attribute VB_Name = "TW"
Option Explicit
Option Private Module
Sub fetchTweetsNewQuery()
    dataSource = "TW"
    Call checkOperatingSystem
    Call setDatasourceVariables
    Call markToCurrentQuery
    Call fetchTweets
End Sub

Sub fetchTweets()
    Dim resultArr As Variant
    Dim col As Integer
    Dim rivi As Integer
    Dim firstButtonLeft As Double
    Dim buttonNum As Integer

    Application.ScreenUpdating = False

    Call protectSheets

    If Range("TWsearchTerm").value = vbNullString Or Range("TWsearchTerm").value = 0 Then
        MsgBox "Twitter search term not set"
        End
    End If

    If usingMacOSX = False Then ProgressBox.Show False
    Call updateProgress(4, "Fetching tweets...", , True)

    resultArr = getTweets(getTokenFromSheet("Twitter"), Range("TWsearchTerm").value, Range("TWcolumns").value, Range("maxResults").value, True, Range("TWresultType").value, Range("TWlanguageCode").value, Range("TWgeoCode").value, Range("TWuntilDate").value, "", Range("TWtimeZone").value, True)

    Call updateProgress(85, "Formatting...", , True)

    sheetName = Range("wsname").value

    If SheetExists(sheetName) = False Then
        Set dataSheet = ThisWorkbook.Sheets.Add
        dataSheet.Name = sheetName
        dataSheet.Tab.ColorIndex = 13
        If Twitter.Visible = xlSheetVisible Then
            dataSheet.move after:=Twitter
        ElseIf YouTube.Visible = xlSheetVisible Then
            dataSheet.move after:=YouTube
        ElseIf Facebook.Visible = xlSheetVisible Then
            dataSheet.move after:=Facebook
        ElseIf BingAds.Visible = xlSheetVisible Then
            dataSheet.move after:=BingAds
        ElseIf AdWords.Visible = xlSheetVisible Then
            dataSheet.move after:=AdWords
        Else
            dataSheet.move after:=Analytics
        End If
        updatingPreviouslyCreatedSheet = False
    Else
        Set dataSheet = ThisWorkbook.Worksheets(sheetName)
        updatingPreviouslyCreatedSheet = True
    End If

    If sendMode = True Then Call checkE(email, dataSource)

    If updatingPreviouslyCreatedSheet = False Then
        sheetID = Range("sheetID").value
    Else
        sheetID = dataSheet.Cells(1, 1).value
        sheetID = findRangeName(dataSheet.Cells(1, 1))
    End If

    Range("queryRunTime").value = Now()


    With dataSheet

        With .Cells(1, 13)
            .Resize(1, UBound(resultArr, 2)).EntireColumn.ClearContents
            .Resize(UBound(resultArr, 1), UBound(resultArr, 2)).value = resultArr
            For col = 1 To UBound(resultArr, 2)
                If .Offset(, col - 1).value = "Followers" Or .Offset(, col - 1).value = "Retweets" Then
                    .Offset(, col - 1).EntireColumn.NumberFormat = "0"
                ElseIf .Offset(, col - 1).value = "Link" Then
                    With dataSheet
                        For rivi = 2 To UBound(resultArr, 1) + 1
                            .Hyperlinks.Add Cells(rivi, 12 + col), Cells(rivi, 12 + col).value
                        Next rivi
                    End With
                End If

                If .Offset(, col - 1).value = "Tweet" Then
                    .Offset(, col - 1).EntireColumn.ColumnWidth = 100
                Else
                    .Offset(, col - 1).EntireColumn.ColumnWidth = 20
                End If

            Next col


            If Not updatingPreviouslyCreatedSheet Then
                With .Resize(1, UBound(resultArr, 2))
                    .Font.Bold = True
                    If Range("doAutofilter").value <> False Then .AutoFilter
                End With
            End If
        End With

        If Not updatingPreviouslyCreatedSheet Then

            .Rows("2:2").Select
            ActiveWindow.FreezePanes = True
            .Cells(1, 13).Select

            .Cells(1, 1).Resize(1, 2).EntireColumn.Hidden = True
            .Select
            .Cells(1, 1).value = sheetID
            .Cells(1, 1).Name = sheetID

            .Cells.Interior.ColorIndex = 2

            With .Cells(2, 4)
                .value = UCase("Twitter report")
                With .Resize(1, 3)
                    .Interior.ColorIndex = 37
                    .Font.ColorIndex = 2
                End With
                .Offset(1).value = "Fetched"
                .Offset(1, 1).value = Now()
                .Offset(1, 2).value = Now()

                .Offset(1, 1).NumberFormatLocal = Range("numformatDate").NumberFormatLocal
                .Offset(1, 2).NumberFormatLocal = Range("numformatTime").NumberFormatLocal

            End With

            .Cells(1, 4).Resize(1, 3).EntireColumn.Font.Size = 9


            firstButtonLeft = Round(.Cells(1, reportStartColumn + 4).Left + buttonSpaceBetween)

            progresspct = Evaluate(progresspct & "+" & "(100-" & progresspct & ")" & "*0.05")
            Call updateProgress(progresspct, "Inserting buttons...")

            Dim createdButtonNum As Integer
            createdButtonNum = 1



            For buttonNum = 1 To 4

                '   Set buttonObj = dataSheet.Shapes.AddTextbox(1, 342, 15, 118, 29)
                Set buttonObj = dataSheet.Shapes.AddShape(5, 10, 10, 200, 40)  '5=msoShapeRoundedRectangle
                With buttonObj

                    .Adjustments(1) = 0.1

                    With .TextFrame
                        .HorizontalAlignment = xlHAlignCenter
                        .VerticalAlignment = xlVAlignCenter
                        .MarginLeft = 0
                        .MarginRight = 0
                        .MarginTop = 0
                        .MarginBottom = 0
                        .Characters.Font.Color = buttonFontColor
                        .Characters.Font.Size = 8
                        .Characters.Font.Name = "Calibri Ligth"
                    End With

                    .Fill.ForeColor.RGB = buttonColour
                    .Line.ForeColor.RGB = buttonBorderColour
                    .Height = buttonHeight
                    .Width = buttonWidth
                    .Top = dataSheet.Cells(2, 1).Top
                    .Left = firstButtonLeft + (createdButtonNum - 1) * (buttonWidth + buttonSpaceBetween)


                    Select Case buttonNum
                    Case 1
                        .OnAction = "refreshDataOnSelectedSheet"
                        .TextFrame.Characters.Text = "REFRESH"
                        .Name = sheetID & "RefreshButton"
                        createdButtonNum = createdButtonNum + 1
                    Case 2
                        .OnAction = "exportReportToExcel"
                        .TextFrame.Characters.Text = "EXPORT TO EXCEL"
                        .Name = sheetID & "ExportExcelButton"
                        createdButtonNum = createdButtonNum + 1
                    Case 3
                        .OnAction = "selectActiveReportInQuerystorage"
                        .TextFrame.Characters.Text = "MODIFY QUERY"
                        .Name = sheetID & "ModifyQueryButton"
                        createdButtonNum = createdButtonNum + 1
                    Case 4
                        .OnAction = "removeSheet"
                        .TextFrame.Characters.Text = "REMOVE SHEET"
                        .Fill.ForeColor.RGB = buttonColourRed
                        .Name = sheetID & "RemoveSheetButton"
                        createdButtonNum = createdButtonNum + 1
                    End Select


                End With

            Next buttonNum
        End If

    End With

    Call copyCurrentquerytoQueryStorage
    Call hideProgressBox
    Call protectSheets

End Sub
