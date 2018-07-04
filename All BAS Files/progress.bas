Attribute VB_Name = "progress"
Option Private Module
Option Explicit

Sub testProgress()
    Call updateProgress(30, "Testing progress box...", "Testing the additional message")

End Sub
Sub stopProcesses()
    On Error Resume Next
    If debugMode Then On Error GoTo 0
    Dim URL As String
    URL = "https://supermetrics.com/api/stopProcess?pid=" & processIDsStr & "&responseFormat=RSCL"
    If Not useQTforDataFetch Then
        Set objHTTPstatus = Nothing
        Call setMSXML(objHTTPstatus)
        If useProxy = True Then objHTTPstatus.setProxy 2, proxyAddress
        objHTTPstatus.Open "GET", URL, False
        If useProxyWithCredentials = True Then objHTTPstatus.setProxyCredentials proxyUsername, proxyPassword
        objHTTPstatus.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        objHTTPstatus.setTimeouts 100000, 100000, 100000, 100000
        objHTTPstatus.setOption 2, 13056
        objHTTPstatus.send ("")
    Else
        Call fetchDataWithQueryTableDirect(URL, "")
    End If

End Sub
Sub getProcessStatus()
    On Error Resume Next
    '  If debugMode Then On Error GoTo 0

    Dim resp As String
    Dim URL As String
    If Not objHTTPstatusRunning Then
        URL = "https://supermetrics.com/api/getQueryStatus?pid=" & processIDsStr & "&type=multi&responseFormat=RSCL&rscL1=" & uriEncode(rscL1)
        Call setMSXML(objHTTPstatus)
        If useProxy = True Then objHTTPstatus.setProxy 2, proxyAddress
        objHTTPstatus.Open "GET", URL, True
        If useProxyWithCredentials = True Then objHTTPstatus.setProxyCredentials proxyUsername, proxyPassword
        objHTTPstatus.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        objHTTPstatus.setTimeouts 1000000, 1000000, 1000000, 1000000
        objHTTPstatus.setOption 2, 13056
        objHTTPstatus.send ("")
        objHTTPstatusRunning = True
    Else
        If objHTTPstatus.readyState = 4 Then
            objHTTPstatusRunning = False
            resp = objHTTPstatus.responsetext
            Set objHTTPstatus = Nothing
            If parseVarFromStr(resp, "QUERIESDONE", rscL1) <> vbNullString Then
                processQueriesCompleted = CLng(parseVarFromStr(resp, "QUERIESDONE", rscL1))
                processQueriesTotal = CLng(parseVarFromStr(resp, "QUERIESTOTAL", rscL1))
            End If
        Else
            objHTTPstatus.waitForResponse 0
        End If
    End If
End Sub
Sub updateProgress(pctDone As Long, Optional currentAction As String, Optional otherLabelText As String, Optional doScreenUpdate As Boolean = True)

    On Error Resume Next
    ' If debugMode = True Then On Error GoTo 0

    If pctDone < pctDonePrev Then
        pctDone = pctDonePrev
    Else
        pctDonePrev = pctDone
    End If

    Dim SBtext As String
    Dim SBtextShape As String

    Dim progressBar As Object
    Dim progressBox2 As Object
    Dim progressBoxAM As Object

    Dim progressIbox1 As Object
    Dim progressIbox2 As Object
    Dim progressIbox3 As Object

    Dim fontColourIndex As Integer
    Dim boxColour As Long


    Dim shTop As Double
    Dim shLeft As Double
    Dim keepOtherLabelText As Boolean
    If otherLabelText = "KEEP" Then
        keepOtherLabelText = True
        otherLabelText = ""
    End If


    Dim progressBoxStopButton As Object

    Application.EnableEvents = False

    loopTimer = Timer

    SBtext = Round(pctDone, 0) & " %  " & currentAction & "   " & otherLabelText
    SBtextShape = vbCrLf & vbCrLf & vbCrLf & currentAction    '& vbCrLf & vbCrLf & otherLabelText


    If usingMacOSX = False Then
        With ProgressBox
            If currentAction <> "" Then .actionLabel.Caption = currentAction
            .FrameProgress.Caption = CStr(Format(pctDone / 100, "0%"))
            .LabelProgress.Width = pctDone * .FrameProgress.Width / 100
            If Not keepOtherLabelText Then .otherLabel.Caption = otherLabelText
        End With
    Else
        boxColour = buttonColour
        fontColourIndex = 1
        If isSheetAconfigSheet(ActiveSheet.Name) Then Call unprotectSheets
        If shapeExists("progressBox2") = False Then

            shTop = ActiveSheet.Cells(ActiveWindow.ScrollRow, 1).Top + 10
            shLeft = ActiveSheet.Cells(1, ActiveWindow.ScrollColumn).Left + 10

            Set progressBox2 = ActiveSheet.Shapes.AddTextbox(1, 342, 15, 118, 29)
            With progressBox2
                .TextFrame.HorizontalAlignment = xlHAlignLeft
                .TextFrame.VerticalAlignment = xlVAlignTop
                .TextFrame.Characters.Text = SBtextShape
                .TextFrame.Characters.Font.ColorIndex = fontColourIndex
                .TextFrame.Characters.Font.Size = 10
                .TextFrame.Characters.Font.Bold = True
                .Fill.ForeColor.RGB = boxColour
                .Line.BackColor.RGB = boxColour
                .Line.ForeColor.RGB = buttonBorderColour
                .Height = 180
                .Width = 300
                .Top = shTop
                .Left = shLeft
                .Name = "progressBox2"
                .ZOrder (0)  'BringToFront
            End With

            Set progressBoxAM = ActiveSheet.Shapes.AddTextbox(1, 342, 15, 118, 29)
            With progressBoxAM
                .TextFrame.HorizontalAlignment = xlHAlignLeft
                .TextFrame.VerticalAlignment = xlVAlignTop
                If Not keepOtherLabelText Then .TextFrame.Characters.Text = otherLabelText
                .TextFrame.Characters.Font.ColorIndex = fontColourIndex
                .TextFrame.Characters.Font.Size = 10
                .TextFrame.Characters.Font.Bold = True
                .Fill.ForeColor.RGB = boxColour
                .Line.BackColor.RGB = boxColour
                .Line.ForeColor.RGB = boxColour
                .Height = 50
                .Width = 300 - 2
                .Top = shTop + 100
                .Left = shLeft + 1
                .Name = "progressBoxAM"
                .ZOrder (0)   '0
            End With

            Set progressBar = ActiveSheet.Shapes.AddTextbox(1, 342, 15, 118, 29)
            With progressBar
                .TextFrame.HorizontalAlignment = xlHAlignLeft
                .TextFrame.VerticalAlignment = xlVAlignCenter
                .TextFrame.Characters.Text = Round(pctDone, 0) & "%"
                .TextFrame.Characters.Font.ColorIndex = 1
                .TextFrame.Characters.Font.Size = 10
                .TextFrame.Characters.Font.Bold = True
                .TextFrame.MarginLeft = 1
                .TextFrame.MarginRight = 0
                .Fill.ForeColor.RGB = RGB(185, 250, 0)
                .Height = 20
                .Left = progressBox2.Left + 8
                .Width = pctDone * (progressBox2.Width - 25) / 100
                .Top = progressBox2.Top + 10
                .Name = "progressBar"
                .ZOrder (0)  '0
            End With


            Set progressIbox1 = ActiveSheet.Shapes.AddTextbox(1, 342, 15, 118, 29)
            With progressIbox1
                .Fill.ForeColor.RGB = RGB(90, 90, 90)
                .Height = 5
                .Left = progressBox2.Left + 8
                .Width = 5
                .Top = progressBox2.Top + progressBox2.Height - 20
                .Name = "progressIbox1"
                .ZOrder (0)  '0
                .Visible = False
            End With

            Set progressIbox2 = ActiveSheet.Shapes.AddTextbox(1, 342, 15, 118, 29)
            With progressIbox2
                .Fill.ForeColor.RGB = RGB(90, 90, 90)
                .Height = 5
                .Left = progressBox2.Left + 8 + 15
                .Width = 5
                .Top = progressBox2.Top + progressBox2.Height - 20
                .Name = "progressIbox2"
                .ZOrder (0)
                .Visible = False
            End With

            Set progressIbox3 = ActiveSheet.Shapes.AddTextbox(1, 342, 15, 118, 29)
            With progressIbox3
                .Fill.ForeColor.RGB = RGB(90, 90, 90)
                .Height = 5
                .Left = progressBox2.Left + 8 + 30
                .Width = 5
                .Top = progressBox2.Top + progressBox2.Height - 20
                .Name = "progressIbox3"
                .ZOrder (0)
                .Visible = False
            End With


            Set progressBoxStopButton = ActiveSheet.Shapes.AddTextbox(1, 342, 15, 118, 29)
            With progressBoxStopButton
                .OnAction = "stopExecution"
                .TextFrame.HorizontalAlignment = xlHAlignCenter
                .TextFrame.VerticalAlignment = xlVAlignCenter
                .TextFrame.Characters.Text = "Stop"
                .TextFrame.Characters.Font.ColorIndex = 1
                .TextFrame.Characters.Font.Size = 12
                .TextFrame.Characters.Font.Bold = True
                .Fill.ForeColor.RGB = buttonColourRed
                .Height = 20
                .Width = 50
                .Top = progressBox2.Top + progressBox2.Height - .Height - 10
                .Left = progressBox2.Left + progressBox2.Width - .Width - 10
                .Name = "progressBoxStopButton"
                .ZOrder (0)
            End With
        Else
            Set progressBox2 = ActiveSheet.Shapes("progressBox2")
            With progressBox2
                .TextFrame.Characters.Text = SBtextShape
                .ZOrder (0)
            End With
            Set progressBoxAM = ActiveSheet.Shapes("progressBoxAM")
            With progressBoxAM
                .TextFrame.Characters.Text = otherLabelText
                .ZOrder (0)
            End With
            Set progressBar = ActiveSheet.Shapes("progressBar")
            With progressBar
                .Width = pctDone * (progressBox2.Width - 10) / 100
                .TextFrame.Characters.Text = Round(pctDone, 0) & " %"
                .ZOrder (0)
            End With

            Set progressIbox1 = ActiveSheet.Shapes("progressIbox1")
            With progressIbox1
                .ZOrder (0)
            End With
            Set progressIbox2 = ActiveSheet.Shapes("progressIbox2")
            With progressIbox2
                .ZOrder (0)
            End With
            Set progressIbox3 = ActiveSheet.Shapes("progressIbox3")
            With progressIbox3
                .ZOrder (0)
            End With

            Set progressBoxStopButton = ActiveSheet.Shapes("progressBoxStopButton")
            progressBoxStopButton.ZOrder (0)

        End If

        If doScreenUpdate = True Then
            Application.ScreenUpdating = True
            Application.ScreenUpdating = False
        End If


    End If


    Application.StatusBar = Left(SBtext, 255)
    DoEvents
End Sub

Sub updateProgressAdditionalMessage(otherLabelText As String)
    On Error Resume Next
    Dim progressBoxAM As Object
    If usingMacOSX = False Then
        ProgressBox.otherLabel.Caption = otherLabelText
        DoEvents
    Else
        Application.ScreenUpdating = True
        Set progressBoxAM = ActiveSheet.Shapes("progressBoxAM")
        With progressBoxAM
            .TextFrame.Characters.Text = otherLabelText
            .ZOrder (0)
        End With

        ActiveSheet.Shapes("progressBoxStopButton").ZOrder (0)

    End If
    DoEvents
    Application.ScreenUpdating = False
End Sub

Sub updateProgressIterationBoxes(Optional action As String = vbNullString, Optional changeInterval As Long = 1)

    On Error Resume Next

    Dim timeDiff As Double
    Dim minTimeDiff As Double

    Dim box1 As Object
    Dim box3 As Object
    Dim box2 As Object

    If usingMacOSX = True Then
        Set box1 = ActiveSheet.Shapes("progressIbox1")
        Set box2 = ActiveSheet.Shapes("progressIbox2")
        Set box3 = ActiveSheet.Shapes("progressIbox3")
        box1.ZOrder (0)
        box2.ZOrder (0)
        box3.ZOrder (0)
    Else
        Set box1 = ProgressBox.prog1
        Set box2 = ProgressBox.prog2
        Set box3 = ProgressBox.prog3
    End If


    If action = "EXITLOOP" Then
        box1.Visible = False
        box2.Visible = False
        box3.Visible = False
        DoEvents
        loopIterationCount = 0
        Exit Sub
    End If

    timeDiff = Timer - loopTimer
    minTimeDiff = 0.25

    If loopIterationCount = 0 Then
        If timeDiff < minTimeDiff Then Exit Sub
        box1.Visible = False
        box2.Visible = False
        box3.Visible = False
        DoEvents
        loopTimer = Timer
        loopIterationCount = loopIterationCount + 1
    ElseIf loopIterationCount = changeInterval Then
        If timeDiff < minTimeDiff Then Exit Sub
        box1.Visible = True
        box2.Visible = False
        box3.Visible = False
        DoEvents
        loopTimer = Timer
        loopIterationCount = loopIterationCount + 1
    ElseIf loopIterationCount = changeInterval * 2 Then
        If timeDiff < minTimeDiff Then Exit Sub
        box1.Visible = True
        box2.Visible = True
        box3.Visible = False
        DoEvents
        loopTimer = Timer
        loopIterationCount = loopIterationCount + 1
    ElseIf loopIterationCount = changeInterval * 3 Then
        If timeDiff < minTimeDiff Then Exit Sub
        box1.Visible = True
        box2.Visible = True
        box3.Visible = True
        DoEvents
        loopTimer = Timer
        loopIterationCount = loopIterationCount + 1
    ElseIf loopIterationCount = changeInterval * 4 Then
        loopIterationCount = 0
        loopTimer = Timer
    Else
        loopIterationCount = loopIterationCount + 1
    End If

End Sub



Sub hideProgressBox()
    On Error Resume Next
    If debugMode = True Then On Error GoTo 0
    Application.StatusBar = False
    Application.ScreenUpdating = False
    Application.Cursor = xlDefault
    Call unprotectSheets

    If usingMacOSX = False Then
        Unload ProgressBox
    Else
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
            If shapeExists("progressBox2", ws.Name) = True Then ws.Shapes("progressBox2").Delete
            If shapeExists("progressBoxAM", ws.Name) = True Then ws.Shapes("progressBoxAM").Delete
            If shapeExists("progressBar", ws.Name) = True Then ws.Shapes("progressBar").Delete
            If shapeExists("progressIbox1", ws.Name) = True Then ws.Shapes("progressIbox1").Delete
            If shapeExists("progressIbox2", ws.Name) = True Then ws.Shapes("progressIbox2").Delete
            If shapeExists("progressIbox3", ws.Name) = True Then ws.Shapes("progressIbox3").Delete
            If shapeExists("progressBoxStopButton", ws.Name) = True Then ws.Shapes("progressBoxStopButton").Delete
        Next
    End If
    pctDonePrev = 0
    Call protectSheets
End Sub

Sub removeTempsheet()
    Application.DisplayAlerts = False
    On Error Resume Next
    If Not tempSheet Is Nothing Then tempSheet.Delete
    If debugMode = True Then On Error GoTo 0
    Application.DisplayAlerts = True
End Sub
Sub removeDatasheet()
    Application.DisplayAlerts = False
    On Error Resume Next
    If Not dataSheet Is Nothing Then dataSheet.Delete
    If debugMode = True Then On Error GoTo 0
    Application.DisplayAlerts = True
End Sub
Sub stopExecution()
    On Error Resume Next
    Call hideProgressBox
    Application.DisplayAlerts = False
    On Error Resume Next
    If Not tempSheet Is Nothing Then tempSheet.Delete
    If runningMultipleReports = True And IsArray(profileSelectionsArr) Then Range("profileSelections" & varsuffix).value = profileSelectionsArr
    If debugMode = True Then On Error GoTo 0
    Application.DisplayAlerts = True
    If inDataFetchLoop Then Call stopProcesses
    Call eraseObjHTTPs
    Application.StatusBar = False
    inDataFetchLoop = False
    pctDonePrev = 0
    End
End Sub

