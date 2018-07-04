Attribute VB_Name = "fileManagement"
Option Private Module
Option Explicit

Sub refreshAndExport()
    refreshBeforeExport = True
    exportUF.Show
End Sub

Sub launchExportUF()
    refreshBeforeExport = False
    exportUF.Show
End Sub

Sub exportReportToExcel(Optional exportAllSheets As Boolean = False, Optional savePath As String = "", Optional closeFile As Boolean = False, Optional thisSheet As Worksheet, Optional appendWB As Workbook)



    Application.ScreenUpdating = False
    Dim i As Long
    If thisSheet Is Nothing Then
        Set thisSheet = ActiveSheet
    End If
    Dim newWB As Workbook
    Set newWB = Workbooks.Add
    Dim newSheet As Worksheet
    Dim newSheetOrig As Worksheet
    Dim ws As Worksheet
    Dim tempSheet As Worksheet
    Dim fileName As String
    Dim fileName2 As String


    If exportAllSheets = False Then
        thisSheet.Copy before:=newWB.Sheets(1)
        Set newSheet = newWB.Worksheets(1)
    Else
        Set tempSheet = newWB.Sheets.Add
        Set newSheet = tempSheet
    End If

    Application.DisplayAlerts = False
    For Each ws In newWB.Worksheets
        If ws.Name <> newSheet.Name Then ws.Delete
    Next
    Application.DisplayAlerts = True

    Call deleteButtonsFromSheet(newSheet)
    With newSheet.Columns("A:B")
        .Hidden = False
        .Delete
    End With


    If exportAllSheets = True Then
        For Each ws In ThisWorkbook.Worksheets
            If isSheetAconfigSheet(ws.Name) = False And ws.Visible = xlSheetVisible Then
                reportsFound = True
                ws.Copy after:=newWB.Sheets(newWB.Sheets.Count)
                Set newSheet = newWB.Worksheets(newWB.Sheets.Count)
                Call deleteButtonsFromSheet(newSheet)
                With newSheet.Columns("A:B")
                    .Hidden = False
                    .Delete
                End With
            End If
        Next
        Application.DisplayAlerts = False
        tempSheet.Delete
        Application.DisplayAlerts = True
    End If

    If Not appendWB Is Nothing Then
        For i = appendWB.Worksheets.Count To 1 Step -1
            appendWB.Worksheets(i).Copy before:=newWB.Sheets(1)
        Next i
    End If

    With newWB
        For i = 1 To 56
            .Colors(i) = ThisWorkbook.Colors(i)
        Next i
    End With

    If savePath <> vbNullString Then
        If Right(savePath, 1) <> "\" And Right(savePath, 1) <> "/" Then
            If usingMacOSX = True Then
                savePath = savePath & "/"
            Else
                savePath = savePath & "\"
            End If
        End If
        If exportAllSheets = True Then
            fileName = "Supermetrics Data Grabber Reports " & Year(Now) & "-" & Format(Month(Now), "00") & "-" & Format(Day(Now), "00")
        Else
            fileName = newSheet.Name & " " & Year(Now) & "-" & Format(Month(Now), "00") & "-" & Format(Day(Now), "00")
        End If
        If FileOrDirExists(savePath & fileName & ".xls") = True Or FileOrDirExists(savePath & fileName & ".xlsx") = True Or FileOrDirExists(savePath & fileName & ".xlsm") = True Then
            For i = 1 To 100
                fileName2 = fileName & " " & i
                If FileOrDirExists(savePath & fileName2 & ".xls") = False And FileOrDirExists(savePath & fileName2 & ".xlsx") = False And FileOrDirExists(savePath & fileName2 & ".xlsm") = False Then
                    fileName = fileName2
                    Exit For
                End If
            Next i
        End If
        newWB.SaveAs savePath & fileName

        If closeFile = True Then
            newWB.Close
        End If
    End If


End Sub

Sub deleteButtonsFromSheet(ws As Worksheet)
    On Error Resume Next
    Dim sh As Object
    Dim shapeName As String
    For Each sh In ws.Shapes
        shapeName = sh.Name
        If InStr(1, shapeName, "RemoveSheetButton") > 0 Or InStr(1, shapeName, "condFormButton") > 0 Or InStr(1, shapeName, "RefreshButton") > 0 Or InStr(1, shapeName, "CreatePPTButton") > 0 Or InStr(1, shapeName, "ChartTypeButton") > 0 Or InStr(1, shapeName, "sortButton") > 0 Or InStr(1, shapeName, "ExportExcelButton") > 0 Or InStr(1, shapeName, "ModifyQueryButton") > 0 Or InStr(1, shapeName, "CTCTB") > 0 Or InStr(1, shapeName, "CASSB") > 0 Or InStr(1, shapeName, "chartCategoriesLabel") > 0 Then
            sh.Delete
        End If
    Next
    For Each sh In ws.DropDowns
        shapeName = sh.Name
        If InStr(1, shapeName, "chartCategoriesDropdown") > 0 Then
            sh.Delete
        ElseIf InStr(1, shapeName, "valuesDropdown") > 0 Then
            sh.Delete
        End If
    Next
End Sub
Sub closeOtherWorkbooks()
    Dim wb As Workbook
    Application.DisplayAlerts = False
    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            wb.Close
        End If
    Next
    Application.DisplayAlerts = True
End Sub

Sub migrateReports()

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim filenameAndPath As Variant
    Dim oldGADG As Object
    Dim newGADG As Object
    Dim oldGADGname As String
    Dim newGADGname As String
    Dim oldGADGpath As String
    Dim oldGADGextension As String
    Dim catLabel As String
    Dim oldGADGqsSheet As Worksheet
    Dim oldGADGwasOpen As Boolean

    Dim lastColOldSheet As Long
    Dim lastColNewSheet As Long

    Dim sht As Worksheet

    Dim reportSheetName As String
    Dim sheetNameRow As Integer
    Dim sheetIDRow As Integer
    Dim sheetNum As Long
    Dim sheetID As String
    Dim sheetObj As Object
    Dim sheetObjNewGADG As Object
    sheetNameRow = Range("querySheetRow").row
    sheetIDRow = Range("querySheetIDRow").row

    Dim foundNonReportSheet As Boolean
    Dim nonReportSheetNames As String
    Dim copyNonReportSheets As VbMsgBoxResult

    Dim versionNumOfOldGADG As Double

    Dim vsarNew As Long
    Dim vsarOld As Long
    Dim vriviOld As Long
    Dim rivi As Long
    Dim rivi2 As Long
    Dim col As Long
    Dim dummy

    Dim showGridlines As Boolean
    showGridlines = True

    Dim showRuler As Boolean
    showRuler = True

    Dim showHeadings As Boolean
    showHeadings = True

    Call checkOperatingSystem
    Application.DisplayAlerts = False
    If usingMacOSX = True Then
        filenameAndPath = Application.GetOpenFilename(Title:="Select old version of " & appName)
    Else
        filenameAndPath = Application.GetOpenFilename("Excel Files ,*.xls;*.xlsm;*.xlsx", , "Select old version of " & appName)
    End If

    If filenameAndPath = False Then End

    oldGADGname = GetFileName(filenameAndPath)
    oldGADGpath = Replace(filenameAndPath, oldGADGname, "")
    oldGADGextension = getExtension(filenameAndPath)

    Set newGADG = ThisWorkbook
    newGADGname = newGADG.Name

    If oldGADGname = newGADGname Then
        FileCopy filenameAndPath, oldGADGpath & "SupermetricsDataGrabber OLD VERSION." & oldGADGextension
        oldGADGname = "SupermetricsDataGrabber OLD VERSION." & oldGADGextension
        filenameAndPath = oldGADGpath & "SupermetricsDataGrabber OLD VERSION." & oldGADGextension
    End If

    If IsBookOpen(oldGADGname) = False Then
        Set oldGADG = Workbooks.Open(filenameAndPath, False, True)
        oldGADGwasOpen = False
    Else
        Set oldGADG = Workbooks(oldGADGname)
        oldGADGwasOpen = True
    End If

    newGADG.Activate

    versionNumOfOldGADG = 0
    If nameExists("versionnum", oldGADG) Then
        versionNumOfOldGADG = val(oldGADG.Sheets("vars").Range("versionnum").value)
    Else
        versionNumOfOldGADG = 0
    End If

    If SheetExists("querystorage", oldGADG) = False Then
        Debug.Print "Querystorage sheet not found"
        MsgBox "No report definitions found in selected file", , "No reports found"
        Application.DisplayAlerts = False
        If oldGADGwasOpen = False Then oldGADG.Close SaveChanges:=False
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        End
    End If

    Set oldGADGqsSheet = oldGADG.Sheets("querystorage")

    With newGADG.Sheets("querystorage")
        vsarNew = findLastCell(newGADG.Sheets("querystorage")).Column
        vsarOld = findLastCell(oldGADGqsSheet).Column
        vriviOld = findLastCell(oldGADGqsSheet).row
        If vsarOld <= 2 Then
            Debug.Print "Report definitions not found in querystorage sheet "
            MsgBox "No report definitions found in selected file", , "No reports found"
            Application.DisplayAlerts = False
            If oldGADGwasOpen = False Then oldGADG.Close SaveChanges:=False
            Application.DisplayAlerts = True
            End
        End If

        'copy definitions
        For rivi = 1 To vikarivi(.Cells(1, 2))
            catLabel = LCase(.Cells(rivi, 2).value)
            If catLabel <> "" Then
                For rivi2 = 1 To vikarivi(oldGADGqsSheet.Cells(1, 2))
                    If LCase(oldGADGqsSheet.Cells(rivi2, 2).value) = catLabel Then
                        If catLabel = "metrics" Then
                            oldGADGqsSheet.Range(oldGADGqsSheet.Cells(rivi2, 3), oldGADGqsSheet.Cells(rivi2 + 13, vsarOld)).Copy .Cells(rivi, vsarNew + 1)
                        ElseIf catLabel = "dimensions" Then
                            oldGADGqsSheet.Range(oldGADGqsSheet.Cells(rivi2, 3), oldGADGqsSheet.Cells(rivi2 + 11, vsarOld)).Copy .Cells(rivi, vsarNew + 1)
                        ElseIf catLabel = "profiles" Then
                            oldGADGqsSheet.Range(oldGADGqsSheet.Cells(rivi2, 3), oldGADGqsSheet.Cells(rivi2 + vriviOld, vsarOld)).Copy .Cells(rivi, vsarNew + 1)
                        ElseIf catLabel = "sheetid" Then    'dont copy sheetid
                        Else
                            oldGADGqsSheet.Range(oldGADGqsSheet.Cells(rivi2, 3), oldGADGqsSheet.Cells(rivi2, vsarOld)).Copy .Cells(rivi, vsarNew + 1)
                        End If
                        Exit For
                    End If
                Next rivi2
            End If
        Next rivi

        'check duplicate sheet names
        For col = vsarNew + 1 To findLastCell(newGADG.Sheets("querystorage")).Column
            reportSheetName = .Cells(sheetNameRow, col).value
            If reportSheetName <> "" Then
                If SheetExists(reportSheetName) = True Then
                    For sheetNum = 1 To 1000
                        If Not SheetExists(reportSheetName & sheetNum) Then
                            .Cells(sheetNameRow, col).value = reportSheetName & sheetNum
                            Exit For
                        End If
                    Next sheetNum
                End If
            End If
        Next col

        'delete definitions that don't exist in old version, update sheet names
        For col = findLastCell(newGADG.Sheets("querystorage")).Column To vsarNew + 1 Step -1
            sheetID = .Cells(sheetIDRow, col).value
            reportSheetName = findSheetNameForSheetID(sheetID, oldGADG)
            If reportSheetName = vbNullString Then
                .Cells(1, col).EntireColumn.Delete
            Else
                If SheetExists(reportSheetName, oldGADG) = False Then
                    .Cells(1, col).EntireColumn.Delete
                Else
                    .Cells(sheetNameRow, col).value = reportSheetName
                End If
            End If
        Next col

        importingFromOldVersion = True

        For col = vsarNew + 1 To findLastCell(newGADG.Sheets("querystorage")).Column


            'get dates
            dataSource = .Cells(Range("datasourceRow").row, col).value

            Call setDatasourceVariables

            sheetID = .Cells(sheetIDRow, col).value
            reportSheetName = findSheetNameForSheetID(sheetID, oldGADG)
            If nameExists(sheetID & "_sdate", oldGADG) And nameExists(sheetID & "_edate", oldGADG) Then
                startDate = oldGADG.Sheets(reportSheetName).Range(sheetID & "_sdate").value
                endDate = oldGADG.Sheets(reportSheetName).Range(sheetID & "_edate").value
                Range("startDate" & varsuffix).value = startDate
                Range("endDate" & varsuffix).value = endDate
            End If

            With Sheets("querystorage").Cells(Range("sortTypeRowQS").row, col)
                If .value = "" Then
                    .value = fetchValue("sortType", oldGADG.Sheets(reportSheetName))
                End If
            End With
            '  Set sheetObjNewGADG = newGADG.Sheets.Add
            'With sheetObjNewGADG
            'oldGADG.Sheets(reportSheetName).Cells.Copy
            '.Name = reportSheetName
            'With .Cells
            '.PasteSpecial Paste:=xlPasteFormulas
            '                        .PasteSpecial Paste:=xlPasteFormats
            '                        .PasteSpecial Paste:=xlPasteColumnWidths
            'End With
            'newGADG.Activate
            '.Select
            'End With
            '
            'Call refreshDataOnSelectedSheet
            reportRunSuccessful = True
            Call runQueryFromQueryStorageCol(col)
            Application.ScreenUpdating = False
            If Not isSheetAconfigSheet(ActiveSheet.Name) Then

                'check if old sheet has calculations, copy to new
                '            If reportRunSuccessful = True Then
                '                'sheetID = .Cells(sheetIDRow, col).value
                '                ' reportSheetName = findSheetNameForSheetID(sheetID, oldGADG)
                '                If reportSheetName <> vbNullString Then
                '                    With oldGADG.Sheets(reportSheetName)
                '                        Debug.Print "Copying left side formulas from old sheet to new.. " & reportSheetName
                '                        If versionNumOfOldGADG < 1.6 Then
                '                            .Range(.Cells(resultStartRow, 3), .Cells(.Range("A1").SpecialCells(xlCellTypeLastCell).Row, resultStartColumn - 3)).Copy
                '                        Else
                '                            .Range(.Cells(resultStartRow, reportStartColumn + 2), .Cells(.Range("A1").SpecialCells(xlCellTypeLastCell).Row, resultStartColumn - 1)).Copy
                '                        End If
                '                        With ActiveSheet.Range(Cells(resultStartRow, reportStartColumn + 2), Cells(.Range("A1").SpecialCells(xlCellTypeLastCell).Row, resultStartColumn - 1))
                '                            .PasteSpecial Paste:=xlPasteFormulas
                '                            .PasteSpecial Paste:=xlPasteFormats
                '                            .PasteSpecial Paste:=xlPasteColumnWidths
                '                        End With
                '                        lastColOldSheet = .Range("A1").SpecialCells(xlCellTypeLastCell).Column
                '                        lastColNewSheet = ActiveSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Column
                '                        If lastColOldSheet > lastColNewSheet Then
                '                            Debug.Print "Copying right side formulas from old sheet to new.. " & reportSheetName
                '                            .Columns(ColumnLetter(lastColNewSheet + 1) & ":" & ColumnLetter(lastColOldSheet)).Copy
                '                            With ActiveSheet.Columns(ColumnLetter(lastColNewSheet + 1))
                '                                .PasteSpecial Paste:=xlPasteFormulas
                '                                .PasteSpecial Paste:=xlPasteFormats
                '                                .PasteSpecial Paste:=xlPasteColumnWidths
                '                            End With
                '                        End If
                '                    End With
                '                End If
                '            End If


                'copy sheet tab colour
                ActiveSheet.Tab.ColorIndex = oldGADG.Sheets(reportSheetName).Tab.ColorIndex
                ActiveSheet.Tab.Color = oldGADG.Sheets(reportSheetName).Tab.Color

                'copy sheet view state properties
                Application.ScreenUpdating = False
                oldGADG.Activate
                Sheets(reportSheetName).Activate
                showGridlines = ActiveWindow.DisplayGridlines
                '     showRuler = ActiveWindow.DisplayRuler
                showHeadings = ActiveWindow.DisplayHeadings
                newGADG.Activate
                ActiveWindow.DisplayGridlines = showGridlines
                ' ActiveWindow.DisplayRuler = showRuler
                ActiveWindow.DisplayHeadings = showHeadings
                Application.ScreenUpdating = True
            End If


        Next col


        'copy non-report sheets
        foundNonReportSheet = False
        For sheetNum = 1 To oldGADG.Sheets.Count
            Set sheetObj = oldGADG.Sheets(sheetNum)
            If sheetObj.Visible = xlSheetVisible And isSheetAconfigSheet(sheetObj.Name) = False Then
                If findSheetIDForSheetName(sheetObj.Name, oldGADG) = vbNullString Then
                    If foundNonReportSheet = False Then
                        nonReportSheetNames = sheetObj.Name
                    Else
                        nonReportSheetNames = nonReportSheetNames & ", " & sheetObj.Name
                    End If
                    foundNonReportSheet = True
                End If
            End If
        Next sheetNum
        If foundNonReportSheet = True Then
            copyNonReportSheets = MsgBox("There are some sheets in the selected file that do not contain a report. Do you wish to copy these to the new version too? These sheets are: " & nonReportSheetNames, vbYesNo, "Copy sheets that don't have reports?")
            If copyNonReportSheets = vbYes Then
                For sheetNum = 1 To oldGADG.Sheets.Count
                    Set sheetObj = oldGADG.Sheets(sheetNum)
                    If sheetObj.Visible = xlSheetVisible And isSheetAconfigSheet(sheetObj.Name) = False Then
                        If findSheetIDForSheetName(sheetObj.Name, oldGADG) = vbNullString Then

                            sheetObj.Copy after:=newGADG.Sheets(newGADG.Sheets.Count)
                            Set sheetObjNewGADG = newGADG.Sheets(newGADG.Sheets.Count)
                            '                            Set sheetObjNewGADG = newGADG.Sheets.Add
                            '                            sheetObjNewGADG.Name = sheetObj.Name
                            '                            sheetObj.Cells.Copy sheetObjNewGADG.Cells
                            Set dummy = Worksheets(1).Range("A1:A1").Find("Dummy", LookIn:=xlValues)
                            sheetObjNewGADG.Cells.Replace What:=oldGADGpath & "[" & oldGADGname & "]", Replacement:="", LookAt:=xlPart, _
                                                          MatchCase:=False, SearchFormat:=False, _
                                                          ReplaceFormat:=False
                            sheetObjNewGADG.Cells.Replace What:="[" & oldGADGname & "]", Replacement:="", LookAt:=xlPart, _
                                                          MatchCase:=False, SearchFormat:=False, _
                                                          ReplaceFormat:=False
                            '                             sheetObjNewGADG.Cells.Replace What:="*[" & oldGADGname & "]", Replacement:="=", LookAt:=xlPart, _
                                                          '                                                          MatchCase:=False, SearchFormat:=False, _
                                                          '                                                          ReplaceFormat:=False




                            '                            'copy sheet tab colour
                            '                            sheetObjNewGADG.Tab.ColorIndex = sheetObj.Tab.ColorIndex
                            '                            sheetObjNewGADG.Tab.Color = sheetObj.Tab.Color
                            '
                            '
                            '                            'copy sheet view state properties
                            '                            Application.ScreenUpdating = False
                            '                            oldGADG.Activate
                            '                            sheetObj.Activate
                            '                            showGridlines = ActiveWindow.DisplayGridlines
                            '                            '    showRuler = ActiveWindow.DisplayRuler
                            '                            showHeadings = ActiveWindow.DisplayHeadings
                            '                            newGADG.Activate
                            '                            sheetObjNewGADG.Select
                            '                            ActiveWindow.DisplayGridlines = showGridlines
                            '                            '    ActiveWindow.DisplayRuler = showRuler
                            '                            ActiveWindow.DisplayHeadings = showHeadings
                            '                            Application.ScreenUpdating = True


                        End If
                    End If
                Next sheetNum
            End If
        End If


        Application.DisplayAlerts = False
        If oldGADGwasOpen = False Then oldGADG.Close SaveChanges:=False
        Application.DisplayAlerts = True

        MsgBox "Copying reports done!", , "Done"

    End With

    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub




Function GetFileName(FullPath)

    Dim StrFind As String
    Dim iCount As Long

    Do Until Left(StrFind, 1) = "\"

        iCount = iCount + 1

        StrFind = Right(FullPath, iCount)

        If iCount = Len(FullPath) Then Exit Do

    Loop
    GetFileName = Right(StrFind, Len(StrFind) - 1)
End Function


Public Function getExtension(fileName) As String

    Dim i As Long
    Dim c As String
    Dim pos As Long

    For i = Len(fileName) To 2 Step -1
        c = Mid(fileName, i, 1)
        If c = "." Then
            pos = i + 1
        End If
    Next

    getExtension = Mid(fileName, pos, (Len(fileName) + 1 - pos))
End Function


Function SheetExists(sheetName As String, Optional wb) As Boolean
' returns TRUE if the sheet exists in the active workbook
    If IsMissing(wb) Then Set wb = ThisWorkbook
    SheetExists = False
    On Error GoTo NoSuchSheet
    If Len(wb.Worksheets(sheetName).Name) > 0 Then
        SheetExists = True
        Exit Function
    End If
NoSuchSheet:
End Function


Public Function nameExists(rangeName, Optional wb As Workbook) As Boolean
    On Error Resume Next
    If wb Is Nothing Then Set wb = ThisWorkbook
    nameExists = False
    nameExists = Len(wb.Names(rangeName).Name) <> 0
End Function



Public Function IsBookOpen(BName As String) As Boolean
    Dim WBk As Workbook
    On Error Resume Next
    'If InStr(1, BName, ".xls", 1) = 0 Then BName = BName & ".xls"
    Set WBk = Workbooks(BName)
    IsBookOpen = (Err = 0) + 0
    Set WBk = Nothing
End Function

Function FileOrDirExists(PathName As String) As Boolean
'Macro Purpose: Function returns TRUE if the specified file
'               or folder exists, false if not.
'PathName     : Supports Windows mapped drives or UNC
'             : Supports Macintosh paths
'File usage   : Provide full file path and extension
'Folder usage : Provide full folder path
'               Accepts with/without trailing "\" (Windows)
'               Accepts with/without trailing ":" (Macintosh)

    Dim iTemp As Integer

    'Ignore errors to allow for error evaluation
    On Error Resume Next
    iTemp = GetAttr(PathName)

    'Check if error exists and set response appropriately
    Select Case Err.Number
    Case Is = 0
        FileOrDirExists = True
    Case Else
        FileOrDirExists = False
    End Select

    'Resume error checking
    On Error GoTo 0
End Function


