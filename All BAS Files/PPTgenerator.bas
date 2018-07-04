Attribute VB_Name = "PPTgenerator"
Option Private Module
Option Explicit

Sub selectPPTtemplate()
    Dim filenameAndPath As Variant
    Call checkOperatingSystem
    If usingMacOSX = True Then
        filenameAndPath = Application.GetOpenFilename(Title:="Select PowerPoint template (.pot, .potx, .potm)")
        If Right(filenameAndPath, 4) <> ".pot" And Right(filenameAndPath, 4) <> "potx" And Right(filenameAndPath, 4) <> "potm" Then
            MsgBox "Selected file is not a valid PowerPoint template. The file type of the template must be .pot, .potx or .potm."
            filenameAndPath = vbNullString
        End If
    Else
        filenameAndPath = Application.GetOpenFilename("PowerPoint templates,*.pot;*.potm;*.potx", , "Select PowerPoint template")
    End If

    PPTtemplatePath = vbNullString
    If filenameAndPath = False Then
        PPTtemplatePath = vbNullString
        Range("PPTtemplate").value = vbNullString
        Exit Sub
    End If

    PPTtemplatePath = filenameAndPath

    If FileOrDirExists(PPTtemplatePath) Then
        PPTtemplatePath = filenameAndPath
        Range("PPTtemplate").value = filenameAndPath
    Else
        PPTtemplatePath = vbNullString
        MsgBox "The PowerPoint template path appears to be incorrect. Will continue without the template. The selected path is: " & filenameAndPath
    End If

End Sub

Sub createPPTofActiveSheet()
    linkPPTbox.Show
End Sub


Sub createPPTofActiveSheetInner(Optional exportAllSheets As Boolean = False, Optional savePath As String = "", Optional closeFile As Boolean = False, Optional saveAsPDF As Boolean = False, Optional chartsOnly As Boolean = False, Optional sheetObj As Worksheet)

    Application.StatusBar = "Creating PPT... (Avoid copying anything to the clipboard while this is running)"

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Call checkOperatingSystem

    Dim sheetID As String

    Dim errorMessageShown As Boolean
    errorMessageShown = False

    'luo powerpoint-esityksen
    Dim powerPointApp As Object

    On Error GoTo errhandlerPPTAPP
    Set powerPointApp = CreateObject("PowerPoint.application")
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0


    powerPointApp.Visible = True

    Dim esitys As Object
    Set esitys = powerPointApp.Presentations.Add(True)

    Dim PPTtemplateFile As String
    PPTtemplateFile = Range("PPTtemplate").value
    If FileOrDirExists(PPTtemplateFile) = True Then
        esitys.applytemplate PPTtemplateFile
    End If

    Dim sivu As Object

    Dim kuvaaja As Object
    Dim i As Long
    Dim textBoxObj As Object

    Dim autoFilterRng As Range

    Dim pptChartsPerSlide As Long
    pptChartsPerSlide = Range("pptChartsPerSlide").value

    Dim chartNumOnSlide As Long
    chartNumOnSlide = 1

    Dim pptPasteTypeChart As String
    pptPasteTypeChart = Range("pptPasteTypeChart").value

    Const spaceBetweenCharts As Long = 10

    Dim slideW As Long
    slideW = esitys.PageSetup.slideWidth

    Dim slideH As Long
    slideH = esitys.PageSetup.SlideHeight

    Dim ws As Worksheet
    Dim vrivi As Long

    Dim wsCount As Long
    Dim wsNum As Long

    Dim fileName As String
    Dim fileName2 As String

    Dim widthOrig As Double
    Const maxScale As Single = 1

    If exportAllSheets = True Then
        wsCount = ThisWorkbook.Sheets.Count
    Else
        wsCount = 1
    End If


    fontName = Range("mainFont").value

    For wsNum = 1 To wsCount

        If Not exportAllSheets Then
            If sheetObj Is Nothing Then
                Set ws = ActiveSheet
            Else
                Set ws = sheetObj
            End If
        Else
            Set ws = ThisWorkbook.Sheets(wsNum)
        End If

        If Not isSheetAconfigSheet(ws.Name) And ws.Visible = xlSheetVisible Then
            reportsFound = True
            With ws
                .Select
                chartNumOnSlide = 1
                sheetID = ws.Cells(1, 1).value
                sheetID = findRangeName(ws.Cells(1, 1))



                If sheetID <> vbNullString Then

                    On Error Resume Next
                    .Shapes(sheetID & "sortButton1").Visible = False
                    .Shapes(sheetID & "sortButton2").Visible = False
                    ' If debugMode = True Then On Error GoTo 0


                    If chartsOnly = False Then

                        Set sivu = esitys.slides.Add(esitys.slides.Count + 1, Layout:=12)

                        'add title
                        Set textBoxObj = sivu.Shapes.AddTextbox(1, 40, 30, 300, 50)
                        With textBoxObj
                            .Top = 5
                            .Height = 45
                            .Width = 710
                            .Left = 10
                            With .TextFrame.TextRange
                                '  .ParagraphFormat.Alignment = 2    ' ppAligncenter
                                .Text = ws.Name
                                .Font.Size = 24
                                .Font.Name = fontName
                                '.Font.Bold = True
                                '  .Font.Color = RGB(90, 90, 90)
                            End With
                        End With

                        If .AutoFilterMode = True Then
                            Set autoFilterRng = .AutoFilter.Range
                            .AutoFilterMode = False
                        Else
                            Set autoFilterRng = Nothing
                        End If



                        '   On Error GoTo pptErrHandler
                        If linkedPPT = True Or Range("pptPasteTypeRange").value = "ppPasteOLEObject" Then
                            If Application.CountIf(.Cells(1, 2).EntireColumn, "You have created a PPT that is linked to this sheet. When deleting this sheet, close that PPT first, otherwise Excel might crash.") = 0 Then
                                vrivi = vikarivi(.Cells(1, 2))
                                If vrivi < 5 Then vrivi = 5
                                .Cells(vrivi + 1, 2).value = "You have created a PPT that is linked to this sheet. When deleting this sheet, close that PPT first, otherwise Excel might crash."
                            End If
                            .Range(sheetID & "_PPTrange").Copy
                            If usingMacOSX = True Then
                                sivu.Shapes.Paste
                            Else
                                sivu.Shapes.PasteSpecial dataType:=8, link:=True
                            End If
                        Else
                            Range(sheetID & "_PPTrange").Copy
                            Range(sheetID & "_PPTrange").CopyPicture
                            If usingMacOSX = True Then
                                sivu.Shapes.Paste
                            Else
                                sivu.Shapes.PasteSpecial dataType:=2
                            End If
                        End If

                        If Not autoFilterRng Is Nothing Then
                            autoFilterRng.AutoFilter
                        End If

                        Set kuvaaja = sivu.Shapes(sivu.Shapes.Count)
                        widthOrig = kuvaaja.Width
                        kuvaaja.Width = slideW - 30
                        If kuvaaja.Width / widthOrig > maxScale Then kuvaaja.Width = widthOrig * maxScale
                        kuvaaja.Top = 50
                        kuvaaja.Left = 10
                        If kuvaaja.Top + kuvaaja.Height > slideH Then kuvaaja.Height = slideH - kuvaaja.Top

                    End If

                    Set kuvaaja = Nothing

                    If .ChartObjects.Count > 0 Then
                        For i = 1 To .ChartObjects.Count

                            Application.StatusBar = "Creating PPT... Chart " & i & "  (Avoid copying anything to the clipboard while this is running)"
                            'Debug.Print "KUV: " & .ChartObjects(i).Name

                            If chartNumOnSlide = 1 Then

                                Set sivu = esitys.slides.Add(esitys.slides.Count + 1, Layout:=12)

                                Set textBoxObj = sivu.Shapes.AddTextbox(1, 40, 30, 300, 50)
                                With textBoxObj
                                    .Top = 5
                                    .Height = 45
                                    .Width = 710
                                    .Left = 10
                                    With .TextFrame.TextRange
                                        '  .ParagraphFormat.Alignment = 2    ' ppAligncenter
                                        .Font.Name = fontName
                                        .Text = ws.Name
                                        .Font.Size = 24
                                        ' .Font.Bold = True
                                        '  .Font.Color = RGB(90, 90, 90)
                                    End With
                                End With

                            End If
                            'ppPasteDefault = 0
                            'ppPasteBitmap = 1
                            'ppPasteEnhancedMetafile = 2
                            'ppPasteMetafilePicture = 3
                            'ppPasteGIF = 4
                            'ppPasteJPG = 5
                            'ppPastePNG = 6
                            'ppPasteText = 7
                            'ppPasteHTML = 8
                            'ppPasteRTF = 9
                            'ppPasteOLEObject = 10
                            'ppPasteShape = 11
                            If linkedPPT = True Then
                                .ChartObjects(i).Activate
                                ActiveChart.ChartArea.Copy
                                If usingMacOSX = True Then
                                    sivu.Shapes.Paste
                                Else
                                    sivu.Shapes.PasteSpecial dataType:=10, link:=True
                                End If

                            Else
                                .ChartObjects(i).Select
                                If usingMacOSX = True Then
                                    .ChartObjects(i).Copy
                                    .ChartObjects(i).CopyPicture
                                    sivu.Shapes.Paste
                                Else
                                    Select Case pptPasteTypeChart
                                    Case "ppPastePNG"
                                        .ChartObjects(i).Copy
                                        sivu.Shapes.PasteSpecial dataType:=6
                                    Case "ppPasteEnhancedMetafile"
                                        .ChartObjects(i).Copy
                                        sivu.Shapes.PasteSpecial dataType:=2
                                    Case "ppPasteGIF"
                                        .ChartObjects(i).Copy
                                        sivu.Shapes.PasteSpecial dataType:=4
                                    Case "ppPasteJPG"
                                        .ChartObjects(i).Copy
                                        sivu.Shapes.PasteSpecial dataType:=5
                                    Case "ppPasteOLEObject"
                                        .ChartObjects(i).Copy
                                        sivu.Shapes.PasteSpecial dataType:=10, link:=False
                                    End Select
                                End If
                            End If

                            Set kuvaaja = sivu.Shapes(sivu.Shapes.Count)
                            If pptChartsPerSlide = 1 Then
                                kuvaaja.Width = Evaluate(slideW & "-60")
                                If kuvaaja.Height > slideH Then kuvaaja.Height = Evaluate(slideH & "-60")
                                kuvaaja.Top = Evaluate("(" & slideH & "-" & Round(kuvaaja.Height) & ")" & "/1.3")
                                kuvaaja.Left = Evaluate("(" & slideW & "-" & Round(kuvaaja.Width) & ")" & "/2")
                            Else
                                kuvaaja.Width = 350
                                kuvaaja.Width = Evaluate("(" & slideW & "-3" & "*" & spaceBetweenCharts & ")" & "/2")
                                If kuvaaja.Height > Evaluate("(" & slideH & "-3" & "*" & spaceBetweenCharts & ")" & "/2") Then kuvaaja.Height = Evaluate("(" & slideH & "-3" & "*" & spaceBetweenCharts & ")" & "/2")
                                If chartNumOnSlide = 1 Or chartNumOnSlide = 2 Then
                                    kuvaaja.Top = 50
                                Else
                                    kuvaaja.Top = 50 + (slideH - 50) / 2 + spaceBetweenCharts
                                End If

                                If kuvaaja.Height > (slideH - 50) / 2 Then kuvaaja.Height = (slideH - 50) / 2
                                If kuvaaja.Height + kuvaaja.Top > slideH Then kuvaaja.Height = slideH - kuvaaja.Top

                                If chartNumOnSlide = 1 Or chartNumOnSlide = 3 Then
                                    kuvaaja.Left = spaceBetweenCharts
                                Else
                                    kuvaaja.Left = slideW / 2 + spaceBetweenCharts
                                End If


                                If chartNumOnSlide < 4 Then
                                    chartNumOnSlide = chartNumOnSlide + 1
                                Else
                                    chartNumOnSlide = 1
                                End If

                            End If
                        Next i
                    End If

                    On Error Resume Next
                    .Shapes(sheetID & "sortButton1").Visible = True
                    .Shapes(sheetID & "sortButton2").Visible = True
                    If debugMode = True Then On Error GoTo 0

                End If

            End With

        End If

    Next wsNum


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
            fileName = ActiveSheet.Name & " " & Year(Now) & "-" & Format(Month(Now), "00") & "-" & Format(Day(Now), "00")
        End If
        If saveAsPDF = True Then
            If FileOrDirExists(savePath & fileName & ".pdf") = True Then
                For i = 1 To 100
                    fileName2 = fileName & " " & i
                    If FileOrDirExists(savePath & fileName2 & ".pdf") = False Then
                        fileName = fileName2
                        Exit For
                    End If
                Next i
            End If
        Else
            If FileOrDirExists(savePath & fileName & ".xls") = True Or FileOrDirExists(savePath & fileName & ".xlsx") = True Or FileOrDirExists(savePath & fileName & ".xlsm") = True Then
                For i = 1 To 100
                    fileName2 = fileName & " " & i
                    If FileOrDirExists(savePath & fileName2 & ".ppt") = False And FileOrDirExists(savePath & fileName2 & ".pptx") = False And FileOrDirExists(savePath & fileName2 & ".pptm") = False Then
                        fileName = fileName2
                        Exit For
                    End If
                Next i
            End If
        End If

        If saveAsPDF = True Then
            On Error GoTo errhandlerPDF
            esitys.SaveAs savePath & fileName, FileFormat:=32
            On Error Resume Next
        Else
            esitys.SaveAs savePath & fileName
        End If

        If closeFile = True Then
            esitys.Close
        End If
    End If

    Application.StatusBar = False

    Exit Sub

pptErrHandler:
    If Err.Number = 18 Then End
    If errorMessageShown = False Then
        MsgBox "Error copying charts to Powerpoint. This is probably caused by another application using the clipboard - try not to copy anything during the time this app creates the PPT"
        errorMessageShown = True
    End If
    Resume Next

errhandlerPDF:
    MsgBox "Your PowerPoint version can't save files as PDF. Update to a newer version to enable PDF exports."
    End


errhandlerPPTAPP:
    MsgBox "Can't connect to PowerPoint. A possible reason is that it's not installed on your machine."
    End

End Sub


