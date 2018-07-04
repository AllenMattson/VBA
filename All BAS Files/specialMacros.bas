Attribute VB_Name = "specialMacros"
Option Private Module
Option Explicit

Sub addToTimer(timerNum As Integer, timerName As String)
    On Error GoTo initializeArr
    Dim i As Integer
    If UBound(timerArr) <> 10 Then

    End If
    On Error Resume Next
    timerArr(timerNum, 1) = timerArr(timerNum, 1) + Timer - timerStart
    If timerName <> vbNullString Then timerArr(timerNum, 2) = timerName
    timerStart = Timer
    Exit Sub
initializeArr:
    ReDim timerArr(1 To 10, 1 To 2)
    For i = 1 To 10
        timerArr(i, 1) = 0
    Next i
    timerStart = Timer
    Resume Next
End Sub
Sub displayTimers()
    Dim i As Integer
    For i = 1 To 10
        Debug.Print "TIMER " & i & " " & timerArr(i, 2) & ": " & timerArr(i, 1)
    Next i
End Sub

Sub setPositions()
    Dim sh As Object
    'shape names
    For Each sh In ActiveSheet.Shapes
        sh.Top = sh.Top
    Next
End Sub

Sub newModule()
    On Error Resume Next
    Dim oldSuffix As String
    Dim newSuffix As String
    Dim newSheet As Worksheet
    Dim fromSheet As Worksheet
    Dim shname As String
    Dim n As Name
    Dim sh2 As Shape
    Dim cBox As Object
    Dim newName As String

    'TO CREATE NEW MODULE
    'Create new sheet, type name below
    'Add suffix to changeShapePositionsInModulesSheet macro
    'Run this macro
    'Create macros for the new data source in dsLaunchers
    'Create field areas in vars sheet
    'In vars col B, copy the range names from the pre data source below to col A, change suffix, then run setAndShowVarsSheetColBrangeNames
    'Manually set linkedcells for first dropdowns, then run alignDropdowns macro
    'Remember to add sheet name to isSheetAconfigSheet



    oldSuffix = "MC"
    newSuffix = "TA"

    Set newSheet = Sheets("TwitterAds")
    Set fromSheet = MailChimp

    Call unprotectSheets
    Call showAllShapes
    Call MoveShapesWithCellsAllConfigSh

    fromSheet.Cells.Copy
    newSheet.Select
    Cells.Select
    ActiveSheet.Paste

    Dim sh As Shape

    'shape sizes
    For Each sh In newSheet.Shapes
        Set sh2 = fromSheet.Shapes(sh.Name)
        With sh
            .Width = sh2.Width
            .Height = sh2.Height
            .Top = sh2.Top
            .Left = sh2.Left
        End With
    Next


    'shape names
    For Each sh In ActiveSheet.Shapes
        shname = sh.Name
        If Right(shname, 2) = oldSuffix Then
            sh.Name = Left(shname, Len(shname) - 2) & newSuffix
        End If
    Next


    'dropdown names, linkedcells & actions, shape actions
    With newSheet

        For Each cBox In .DropDowns
            cBox.Name = Replace(cBox.Name, oldSuffix, newSuffix)
            cBox.ListFillRange = Replace(cBox.ListFillRange, oldSuffix, newSuffix)
            cBox.LinkedCell = Replace(cBox.LinkedCell, oldSuffix, newSuffix)
            cBox.OnAction = Replace(cBox.OnAction, oldSuffix, newSuffix)
        Next


        For Each cBox In .Shapes
            cBox.Name = Replace(cBox.Name, oldSuffix, newSuffix)
            If cBox.OnAction <> vbNullString Then cBox.OnAction = Replace(cBox.OnAction, oldSuffix, newSuffix)
        Next

        Dim oldSuffixL As String
        Dim newSuffixL As String
        oldSuffixL = LCase(oldSuffix)
        newSuffixL = LCase(newSuffix)


        For Each cBox In .DropDowns
            cBox.Name = Replace(cBox.Name, oldSuffixL, newSuffixL)
            cBox.ListFillRange = Replace(cBox.ListFillRange, oldSuffixL, newSuffixL)
            cBox.LinkedCell = Replace(cBox.LinkedCell, oldSuffixL, newSuffixL)
            cBox.OnAction = Replace(cBox.OnAction, oldSuffixL, newSuffixL)

        Next


        For Each cBox In .Shapes
            cBox.Name = Replace(cBox.Name, oldSuffix, newSuffix)
            If cBox.OnAction <> vbNullString Then cBox.OnAction = Replace(cBox.OnAction, oldSuffixL, newSuffixL)
        Next
    End With


    'copy range names
    For Each n In ThisWorkbook.Names
        If n.RefersToRange.Worksheet.Name = fromSheet.Name Then
            shname = n.Name
            If Right(shname, 2) = oldSuffix Then
                newName = Left(shname, Len(shname) - 2) & newSuffix
                ActiveWorkbook.Names.Add newName, newSheet.Range(n.RefersToRange.Address)
            Else
                ActiveWorkbook.Names.Add shname, newSheet.Range(n.RefersToRange.Address)
            End If
        End If
    Next n


    'change range name suffix
    For Each n In ActiveWorkbook.Names
        If n.RefersToRange.Worksheet.Name = sheetName Then
            shname = n.Name
            If Right(shname, 2) = oldSuffix Then
                n.Name = Left(shname, Len(shname) - 2) & newSuffix
            End If
        End If
    Next n



    'add buttons to Modules sheet
    Call newLoginButtonsModulesSheet(newSuffix, oldSuffix)




    'position elements on Modules sheet (add new suffix to this macro first)
    Call changeShapePositionsInModulesSheet



End Sub


Sub changeddsuffix()

        Dim oldSuffix As String
        Dim newSuffix As String
           Dim oldSuffixL As String
        Dim newSuffixL As String
        Dim cBox As Object
        
    oldSuffix = "AW"
    newSuffix = "TA"
        
        oldSuffixL = LCase(oldSuffix)
        newSuffixL = LCase(newSuffix)


        For Each cBox In ActiveSheet.DropDowns
            cBox.Name = Replace(cBox.Name, oldSuffixL, newSuffixL)
            cBox.ListFillRange = Replace(cBox.ListFillRange, oldSuffixL, newSuffixL)
            cBox.LinkedCell = Replace(cBox.LinkedCell, oldSuffixL, newSuffixL)
            cBox.OnAction = Replace(cBox.OnAction, oldSuffixL, newSuffixL)

        Next
End Sub

Sub hideFLfromModulesSheet()
'    Dim s As Shape
'    Dim cBox As Object
'
'    With Modules
'
'        For Each s In .Shapes
'            If Right(s.Name, 2) = "FL" Then s.Delete ' s.Visible = False
'        Next
'
'    End With
End Sub
Sub hideGWfromModulesSheet()
    Dim s As Shape
    Dim cBox As Object

    With Modules

        For Each s In .Shapes
            If Right(s.Name, 2) = "GW" Then s.Visible = False
        Next

    End With
End Sub
Sub setReportOptionsPositions()
    On Error Resume Next
    Dim i As Integer
    Dim objName As String

    With ActiveSheet
        For i = 1 To 14
            Select Case i
            Case 1
                objName = "reportOptionsLabel"
            Case 2
                objName = "rawDataReportOB"
            Case 3
                objName = "formattedReportOB"
            Case 4
                objName = "rawDataReportLabel"
            Case 5
                objName = "formattedReportLabel"
            Case 6
                objName = "groupingLabel"
            Case 7
                objName = "rowLimitLabel"
            Case 8
                objName = "groupingDD"
            Case 9
                objName = "rowLimitDD"
            Case 10
                objName = "createChartsCB"
            Case 11
                objName = "createChartsLabel"
            Case 12
                objName = "condFormDropDown"
            Case 13
                objName = "condFormLabel"
            End Select
            .Shapes(objName).Top = Analytics.Shapes(objName).Top
            .Shapes(objName).Left = Analytics.Shapes(objName).Left
        Next i
    End With

End Sub

Sub modifyProfilelistButtons()
    Dim i As Long
    Dim sh As Shape
    Call unprotectSheets
    For i = 1 To 3
        If i = 1 Then
            Set sh = ActiveSheet.Shapes("clearProfileSelectionsButton")
        ElseIf i = 2 Then
            Set sh = ActiveSheet.Shapes("selectAllProfilesButton")
        Else
            Set sh = ActiveSheet.Shapes("refreshProfileListButton")
        End If

        With sh
            .Width = 83
            .Height = 20
            .TextFrame.Characters.Font.Size = 10
            .TextFrame.Characters.Font.Bold = True
            If i < 3 Then
                .Left = ActiveSheet.Shapes("refreshProfileListButton").Left
                .Top = .Height + ActiveSheet.Shapes("refreshProfileListButton").Top + 4
            End If
        End With
    Next i
End Sub


Sub showSheets()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Sheets
        sh.Visible = xlSheetVisible
    Next
End Sub


Sub setAndShowVarsSheetColBrangeNames()
'set names from col A, show names in col B
On Error Resume Next
Dim c As Range
    For Each c In vars.Range("C1:C1000")
    If c.Offset(0, -2).value <> "" Then
 
    ActiveWorkbook.Names.Add _
        Name:=c.Offset(0, -2).value, _
        RefersTo:=c
    End If
    c.Offset(0, -1).value = ""
        Debug.Print c.Name.Name
        c.Offset(0, -1).value = c.Name.Name
    Next
End Sub

Sub showNames2()
    Dim n As Name
    For Each n In ThisWorkbook.Names
        n.Visible = True
    Next

End Sub

Sub showAllShapes()

    Dim s As Shape
    Dim cBox As Object

    With ActiveSheet

        For Each s In .Shapes
            s.Visible = True
        Next

        For Each cBox In .DropDowns
            cBox.Visible = True
        Next

    End With
End Sub
Sub setrowheight()
    Dim i As Long
    Dim sh As Worksheet
    Call unprotectSheets
    For i = 1 To 4
        Select Case i
        Case 1
            Set sh = Analytics
        Case 1
            Set sh = AdWords
        Case 1
            Set sh = BingAds
        Case 1
            Set sh = Facebook
        End Select
        With sh
            .Rows("11:10000").RowHeight = .Rows("11:11").RowHeight
        End With
        Call protectSheets
    Next i
End Sub
Sub deleteSheets()

    Dim sh As Worksheet
    Application.DisplayAlerts = False

    For Each sh In ThisWorkbook.Worksheets
        If Left(sh.Name, Len("Sheet")) = "Sheet" Then sh.Delete
    Next

    For Each sh In ThisWorkbook.Worksheets
        If Left(sh.Name, Len("temp_")) = "temp_" Then sh.Delete
    Next

    For Each sh In ThisWorkbook.Worksheets
        If Left(sh.Name, Len("report")) = "report" Then
            sh.Visible = xlSheetVisible
            sh.Delete
        End If
    Next

End Sub

Sub changeNumberFormatsToGeneral()

    Dim sh As Worksheet
    Application.DisplayAlerts = False

    For Each sh In ThisWorkbook.Worksheets
        If sh.Cells(1, 1).NumberFormat = "[$-F400]h:mm:ss AM/PM" Then sh.Cells.NumberFormat = ""
    Next

End Sub



Sub copyDDmacros()
    On Error Resume Next
    Dim sh As Shape
    For Each sh In ActiveSheet.Shapes
        sh.OnAction = Workbooks("SupermetricsDataGrabber.xlsm").Sheets(ActiveSheet.Name).Shapes(sh.Name).OnAction
        Debug.Print Workbooks("SupermetricsDataGrabber.xlsm").Sheets(ActiveSheet.Name).Shapes(sh.Name).OnAction
    Next
End Sub


Sub removeMacroFileRefs()
    Dim newSheet As Worksheet
    Dim sh As Shape
    For Each newSheet In ThisWorkbook.Sheets
        For Each sh In newSheet.Shapes
            If sh.OnAction <> vbNullString Then
                If InStr(1, sh.OnAction, "!") > 0 Then
                    sh.OnAction = Right(sh.OnAction, Len(sh.OnAction) - InStr(1, sh.OnAction, "!"))
                End If
            End If
        Next
    Next
End Sub

Sub deleteDuplicateSheets()
    Dim sh As Worksheet
    Application.DisplayAlerts = False
    For Each sh In ThisWorkbook.Worksheets
        If Right(sh.Name, 3) = "(2)" Then
            sh.Delete
        End If
    Next
    Application.DisplayAlerts = True
End Sub

Sub copyShapes()
    Dim sh As Shape
    Dim sh2 As Shape
    For Each sh In Analytics.Shapes
        sh.Copy
        Sheets("YouTube").Select
        ActiveSheet.Paste
        Set sh2 = Sheets("YouTube").Shapes(Sheets("YouTube").Shapes.Count)
        sh2.Top = sh.Top
        sh2.Left = sh.Left
    Next
End Sub
Sub printShapeNames()
    Dim sh As Shape
    Dim sh2 As Shape
    For Each sh In ActiveSheet.Shapes
        Debug.Print sh.Name
    Next
End Sub

Sub showAllModuleSheets()
    Analytics.Visible = xlSheetVisible
    AdWords.Visible = xlSheetVisible
    BingAds.Visible = xlSheetVisible
    YouTube.Visible = xlSheetVisible
    Facebook.Visible = xlSheetVisible
End Sub
Sub setShapeVarsuffix()
    Dim sh As Object
    Dim sh2 As Object
    Dim Sheet2 As Worksheet
    Set Sheet2 = Sheets("YouTube")

    Dim oldSuffix As String
    Dim newSuffix As String
    Dim oldname As String
    oldSuffix = "AW"
    newSuffix = "GW"

    For Each sh In ActiveSheet.Shapes
        If Right(sh.Name, 2) = oldSuffix Then
            Debug.Print sh.Name
            oldname = Left(sh.Name, Len(sh.Name) - 2)
            Sheet2.Shapes(oldname).Name = oldname & newSuffix
        End If
    Next

    For Each sh In ActiveSheet.DropDowns
        sh.Name = Replace(sh.Name, "drm" & LCase(oldSuffix), "drm" & LCase(newSuffix))
        sh.Name = Replace(sh.Name, "drsd" & LCase(oldSuffix), "drsd" & LCase(newSuffix))
        sh.Name = Replace(sh.Name, "drd" & LCase(oldSuffix), "drd" & LCase(newSuffix))
    Next
End Sub

Sub findshape()
    Dim sh As Shape
    Dim shname As String
    For Each sh In ActiveSheet.Shapes
        shname = sh.Name
        If shname = "_cal2AC" Then sh.Delete
    Next
End Sub

Sub alignDropdowns()
    Dim i As Long
    Dim varsheetname As String
    Dim origname As String
    On Error Resume Next
    varsuffix = "TA"
    varsheetname = "vars"
    For i = 2 To 12
        ActiveSheet.DropDowns("drm" & varsuffix & "_" & i).Width = ActiveSheet.DropDowns("drm" & varsuffix & "_1").Width
        ActiveSheet.DropDowns("drm" & varsuffix & "_" & i).Height = ActiveSheet.DropDowns("drm" & varsuffix & "_1").Height
        ActiveSheet.DropDowns("drm" & varsuffix & "_" & i).Left = ActiveSheet.DropDowns("drm" & varsuffix & "_1").Left
        ActiveSheet.DropDowns("drm" & varsuffix & "_" & i).Top = ActiveSheet.DropDowns("drm" & varsuffix & "_1").Top + (4 + ActiveSheet.DropDowns("drm" & varsuffix & "_1").Height) * (i - 1)
        ActiveSheet.DropDowns("drm" & varsuffix & "_" & i).ListFillRange = ActiveSheet.DropDowns("drm" & varsuffix & "_1").ListFillRange
        ActiveSheet.DropDowns("drm" & varsuffix & "_" & i).DropDownLines = ActiveSheet.DropDowns("drm" & varsuffix & "_1").DropDownLines
        ActiveSheet.DropDowns("drm" & varsuffix & "_" & i).LinkedCell = varsheetname & "!" & Range(ActiveSheet.DropDowns("drm" & varsuffix & "_1").LinkedCell).Offset(i - 1).Address

    Next i

    ActiveSheet.DropDowns("drd" & varsuffix & "_1").Height = ActiveSheet.DropDowns("drm" & varsuffix & "_1").Height
    ActiveSheet.DropDowns("drsd" & varsuffix & "_1").Height = ActiveSheet.DropDowns("drm" & varsuffix & "_1").Height
    ActiveSheet.DropDowns("drd" & varsuffix & "_1").Width = ActiveSheet.DropDowns("drm" & varsuffix & "_1").Width
    ActiveSheet.DropDowns("drsd" & varsuffix & "_1").Width = ActiveSheet.DropDowns("drm" & varsuffix & "_1").Width

    For i = 2 To 10
        ActiveSheet.DropDowns("drd" & varsuffix & "_" & i).Width = ActiveSheet.DropDowns("drd" & varsuffix & "_1").Width
        ActiveSheet.DropDowns("drd" & varsuffix & "_" & i).Height = ActiveSheet.DropDowns("drd" & varsuffix & "_1").Height
        ActiveSheet.DropDowns("drd" & varsuffix & "_" & i).Left = ActiveSheet.DropDowns("drd" & varsuffix & "_1").Left
        ActiveSheet.DropDowns("drd" & varsuffix & "_" & i).Top = ActiveSheet.DropDowns("drd" & varsuffix & "_1").Top + (4 + ActiveSheet.DropDowns("drd" & varsuffix & "_1").Height) * (i - 1)
        ActiveSheet.DropDowns("drd" & varsuffix & "_" & i).ListFillRange = ActiveSheet.DropDowns("drd" & varsuffix & "_1").ListFillRange
        ActiveSheet.DropDowns("drd" & varsuffix & "_" & i).DropDownLines = ActiveSheet.DropDowns("drd" & varsuffix & "_1").DropDownLines
        ActiveSheet.DropDowns("drd" & varsuffix & "_" & i).LinkedCell = varsheetname & "!" & Range(ActiveSheet.DropDowns("drd" & varsuffix & "_1").LinkedCell).Offset(i - 1).Address
    Next i
    For i = 2 To 2
        ActiveSheet.DropDowns("drsd" & varsuffix & "_" & i).Width = ActiveSheet.DropDowns("drsd" & varsuffix & "_1").Width
        ActiveSheet.DropDowns("drsd" & varsuffix & "_" & i).Height = ActiveSheet.DropDowns("drsd" & varsuffix & "_1").Height
        ActiveSheet.DropDowns("drsd" & varsuffix & "_" & i).Left = ActiveSheet.DropDowns("drsd" & varsuffix & "_1").Left
        ActiveSheet.DropDowns("drsd" & varsuffix & "_" & i).Top = ActiveSheet.DropDowns("drsd" & varsuffix & "_1").Top + (4 + ActiveSheet.DropDowns("drsd" & varsuffix & "_1").Height) * (i - 1)
        ActiveSheet.DropDowns("drsd" & varsuffix & "_" & i).ListFillRange = ActiveSheet.DropDowns("drsd" & varsuffix & "_1").ListFillRange
        ActiveSheet.DropDowns("drsd" & varsuffix & "_" & i).DropDownLines = ActiveSheet.DropDowns("drsd" & varsuffix & "_1").DropDownLines
        ActiveSheet.DropDowns("drsd" & varsuffix & "_" & i).LinkedCell = varsheetname & "!" & Range(ActiveSheet.DropDowns("drsd" & varsuffix & "_1").LinkedCell).Offset(i - 1).Address
    Next i

    'fix dropdown name bug
    For i = 2 To 12
        With ActiveSheet.DropDowns("drm" & varsuffix & "_" & i)
            origname = .Name
            .Name = .Name & "x"
            .Name = origname
        End With
        With ActiveSheet.DropDowns("drd" & varsuffix & "_" & i)
            origname = .Name
            .Name = .Name & "x"
            .Name = origname
        End With
        With ActiveSheet.DropDowns("drsd" & varsuffix & "_" & i)
            origname = .Name
            .Name = .Name & "x"
            .Name = origname
        End With
    Next i

    Dim cBox As Object

    'use same height for all dropdowns
    For Each cBox In ActiveSheet.DropDowns
        cBox.Height = ActiveSheet.DropDowns("drm" & varsuffix & "_1").Height
    Next

    ActiveSheet.DropDowns("comparisonValueTypeDD").Width = ActiveSheet.DropDowns("compareDatesDD").Width
    ActiveSheet.DropDowns("groupingDD").Width = ActiveSheet.DropDowns("condFormDropDown").Width


End Sub

Sub changeDropdownFormats()
    Dim cBox As Object
    With ActiveSheet
        For Each cBox In .DropDowns
            cBox.Display3DShading = True
        Next
    End With
End Sub

Sub deleteFieldDropdowns()
    Dim cBox As Object
    With ActiveSheet
        For Each cBox In .DropDowns
            If Left(cBox.Name, 2) = "dr" Then cBox.Delete
        Next
    End With
End Sub
Sub copyFieldDropdowns()
    Dim cBox As Object
    With Sheets("Facebook")
        For Each cBox In .DropDowns
            If Left(cBox.Name, 2) = "dr" Then
                cBox.Copy
                Sheets("YouTube").Paste
            End If
        Next
    End With
End Sub

Sub printMacrosAssociatedWithShapes()
    Dim cBox As Object
    For Each cBox In ActiveSheet.Shapes
        Debug.Print cBox.OnAction
    Next


End Sub


Sub configureDropdownsForNewModule()

    Dim oldSuffix As String
    Dim newSuffix As String
    Dim cBox As Object
    oldSuffix = "AW"
    newSuffix = "FL"

    With ActiveSheet

        For Each cBox In .DropDowns
            cBox.Name = Replace(cBox.Name, oldSuffix, newSuffix)
            cBox.ListFillRange = Replace(cBox.ListFillRange, oldSuffix, newSuffix)
            cBox.LinkedCell = Replace(cBox.LinkedCell, oldSuffix, newSuffix)
            cBox.OnAction = Replace(cBox.OnAction, oldSuffix, newSuffix)
        Next


        For Each cBox In .Shapes
            cBox.Name = Replace(cBox.Name, oldSuffix, newSuffix)
            If cBox.OnAction <> vbNullString Then cBox.OnAction = Replace(cBox.OnAction, oldSuffix, newSuffix)
        Next


    End With


    oldSuffix = LCase(oldSuffix)
    newSuffix = LCase(newSuffix)

    With ActiveSheet

        For Each cBox In .DropDowns
            cBox.Name = Replace(cBox.Name, oldSuffix, newSuffix)
            cBox.ListFillRange = Replace(cBox.ListFillRange, oldSuffix, newSuffix)
            cBox.LinkedCell = Replace(cBox.LinkedCell, oldSuffix, newSuffix)
            cBox.OnAction = Replace(cBox.OnAction, oldSuffix, newSuffix)

        Next


        For Each cBox In .Shapes
            cBox.Name = Replace(cBox.Name, oldSuffix, newSuffix)
            If cBox.OnAction <> vbNullString Then cBox.OnAction = Replace(cBox.OnAction, oldSuffix, newSuffix)

        Next
    End With



End Sub


Sub copyRangeNames()

    Dim sheetName As String
    Dim newSheetName As String

    Dim oldSuffix As String
    Dim newSuffix As String

    Dim n As Name
    Dim shname As String
    Dim newName As String
    Dim nn As Name

    sheetName = "varsAW"
    newSheetName = "vars"

    oldSuffix = "AW"
    newSuffix = "FL"


    For Each n In ActiveWorkbook.Names
        If n.RefersToRange.Worksheet.Name = sheetName Then

            shname = n.Name
            If Right(shname, 2) = oldSuffix Then
                newName = Left(shname, Len(shname) - 2) & newSuffix
                Set nn = ActiveWorkbook.Names.Add(newName, Sheets(newSheetName).Range(n.RefersToRange.Address))
            End If

        End If
    Next n

End Sub



Sub copyformulas()
    Dim rivi As Long
    For rivi = 1 To 100
        Cells(rivi, 9).Formula = Cells(rivi, 8).Formula
    Next rivi
End Sub

Sub demoNote_Click()
    ActiveWorkbook.FollowHyperlink Address:=ActiveSheet.Shapes("demoNote").Hyperlink.Address, NewWindow:=True
End Sub
Sub copyShapeFormatsInModulesSheet()
    Dim i As Integer
    Dim j As Integer
    Dim sn As String
    Dim soGA As Object
    Dim so As Object
    With Sheets("Modules")
        For i = 2 To 6

            Select Case i
            Case 1
                varsuffix = ""
                dataSource = "GA"
            Case 2
                varsuffix = "AW"
                dataSource = "AW"
            Case 3
                varsuffix = "FB"
                dataSource = "FB"
            Case 4
                varsuffix = "AC"
                dataSource = "AC"
            Case 5
                varsuffix = "YT"
                dataSource = "YT"
            Case 6
                varsuffix = "FL"
                dataSource = "FL"
            End Select

            For j = 1 To 12


                Select Case j
                Case 1
                    sn = "logoutbutton"
                    .Shapes(sn & varsuffix).TextFrame.Characters.Text = .Shapes(sn).TextFrame.Characters.Text

                Case 2
                    sn = "manageLoginsButton"
                    .Shapes(sn & varsuffix).TextFrame.Characters.Text = .Shapes(sn).TextFrame.Characters.Text

                Case 3
                    sn = "addLoginButton"
                Case 4
                    sn = "addLoginButtonNote1"

                Case 5
                    sn = "addLoginButtonNote2"
                    .Shapes(sn & varsuffix).TextFrame.Characters.Text = .Shapes(sn).TextFrame.Characters.Text

                Case 6
                    sn = "bigBox"
                Case 7
                    sn = "loginBoxNote"
                Case 8
                    sn = "loginButtonArrow"
                Case 9
                    sn = "loginButton"
                Case 10
                    sn = "authStatusBox"
                Case 11
                    sn = "licenseNote"
                Case 12
                    sn = "logo"
                End Select


                On Error Resume Next
                .Shapes(sn).PickUp
                .Shapes(sn & varsuffix).Apply
                .Shapes(sn & varsuffix).Height = .Shapes(sn).Height
                .Shapes(sn & varsuffix).Width = .Shapes(sn).Width
                .Shapes(sn & varsuffix).Adjustments(1) = .Shapes(sn).Adjustments(1)
            Next j
        Next i
    End With
End Sub
Sub testShapeTypes()
    Dim i As Integer
    Dim buttonObj As Object
    On Error Resume Next
    For i = 1 To 200
        Set buttonObj = ActiveSheet.Shapes.AddShape(i, 10, 10, 200, 40)
        buttonObj.Top = 50 * i
        buttonObj.Name = "OBJ_" & i
    Next i
End Sub
Sub copyActionButtonFormats()
    Dim i As Integer
    Dim j As Integer
    Dim sn As String
    Dim soGA As Object
    Dim so As Object
    With ActiveSheet
        .Shapes("refreshButton").PickUp
        For j = 1 To 4
            Select Case j
            Case 1
                sn = "refreshButton"


            Case 2
                sn = "exportButton"

            Case 3
                sn = "deleteButton"
            Case 4
                sn = "migrateButton"


            End Select



            .Shapes(sn).Top = .Shapes("refreshButton").Top + (.Shapes(sn).Height + 5) * (j - 1)
            .Shapes(sn).Left = .Shapes("refreshButton").Left
            .Shapes(sn).Apply
            .Shapes(sn).Adjustments(1) = .Shapes("migrateButton").Adjustments(1)
            .Shapes(sn).Height = .Shapes("migrateButton").Height
            .Shapes(sn).Width = .Shapes("migrateButton").Width

        Next j
    End With
End Sub
Sub copyButtonSizes()
    Dim fromObj As Object
    Dim j As Integer
    Dim sn As String
    With ActiveSheet
        Set fromObj = .Shapes("selectAllProfilesButton")
        fromObj.PickUp
        For j = 1 To 4
            Select Case j
            Case 1
                sn = "clearProfileSelectionsButton"
                .Shapes(sn).Top = fromObj.Top
                .Shapes(sn).Left = fromObj.Left
            Case 2
                sn = "refreshProfileListButton"
                .Shapes(sn).Top = fromObj.Top
                .Shapes(sn).Left = fromObj.Left + fromObj.Width + 3
            Case 3
                sn = "clearFieldsButton"
                .Shapes(sn).Top = fromObj.Top
            Case 4
                sn = "clearFiltersButton"
            End Select
            .Shapes(sn).Width = fromObj.Width
            .Shapes(sn).Height = fromObj.Height
            .Shapes(sn).Apply

        Next j
    End With
End Sub
Sub formatShapes()
    Dim i As Integer
    For i = 1 To 4

        Select Case i
        Case 1
            varsuffix = ""
            dataSource = "GA"
        Case 2
            varsuffix = "AW"
            dataSource = "AW"
        Case 3
            varsuffix = "FB"
            dataSource = "FB"
        Case 4
            varsuffix = "AC"
            dataSource = "AC"
        End Select
        Call setDatasourceVariables

        With Modules
            '    .Shapes("authstatusbox" & varsuffix).TextFrame.Characters.Font.ColorIndex = 1
            '  .Shapes("licensenote" & varsuffix).TextFrame.Characters.Font.ColorIndex = 1
            .Shapes("logoutbutton" & varsuffix).Fill.ForeColor.RGB = buttonColourRed
            .Shapes("logoutbutton" & varsuffix).TextFrame.Characters.Font.ColorIndex = 1

            .Shapes("manageLoginsButton" & varsuffix).Fill.ForeColor.RGB = buttonColour
            .Shapes("manageLoginsButton" & varsuffix).TextFrame.Characters.Font.ColorIndex = 1
            .Shapes("addLoginButton" & varsuffix).Fill.ForeColor.RGB = buttonColour
            .Shapes("addLoginButton" & varsuffix).TextFrame.Characters.Font.ColorIndex = 1

            '  .Shapes("addLoginButtonNote1" & varsuffix).Fill.ForeColor.RGB = buttonColour
            .Shapes("addLoginButtonNote1" & varsuffix).TextFrame.Characters.Font.ColorIndex = 1
            .Shapes("addLoginButtonNote2" & varsuffix).TextFrame.Characters.Font.ColorIndex = 1

            .Shapes("addLoginButton" & varsuffix).Left = .Shapes("addLoginButton" & varsuffix).Left - 5
            .Shapes("addLoginButtonNote1" & varsuffix).Left = .Shapes("addLoginButtonNote1" & varsuffix).Left - 5
            .Shapes("addLoginButtonNote2" & varsuffix).Left = .Shapes("addLoginButtonNote2" & varsuffix).Left - 5


            .Shapes("logoutbutton" & varsuffix).Line.ForeColor.RGB = buttonBorderColour
            .Shapes("logoutbutton" & varsuffix).Line.Visible = True
            .Shapes("logoutbutton" & varsuffix).Line.weight = 0.75

            .Shapes("manageLoginsButton" & varsuffix).Line.ForeColor.RGB = buttonBorderColour
            .Shapes("manageLoginsButton" & varsuffix).Line.Visible = True
            .Shapes("manageLoginsButton" & varsuffix).Line.weight = 0.75


            .Shapes("addLoginButton" & varsuffix).Line.ForeColor.RGB = buttonBorderColour
            .Shapes("addLoginButton" & varsuffix).Line.Visible = True
            .Shapes("addLoginButton" & varsuffix).Line.weight = 0.75
        End With
        '        With configsheet.Shapes("clearProfileSelectionsButton")
        '            configsheet.Shapes("selectAllProfilesButton").Left = .Left
        '            configsheet.Shapes("selectAllProfilesButton").Top = .Top
        '            configsheet.Shapes("selectAllProfilesButton").Width = .Width
        '            configsheet.Shapes("selectAllProfilesButton").Height = .Height
        '        End With

    Next i

End Sub
Sub formatProfileSelectionButtons()
    Dim i As Integer
    For i = 1 To 4

        Select Case i
        Case 1
            varsuffix = ""
            dataSource = "GA"
        Case 2
            varsuffix = "AW"
            dataSource = "AW"
        Case 3
            varsuffix = "FB"
            dataSource = "FB"
        Case 4
            varsuffix = "AC"
            dataSource = "AC"
        Case 5
            varsuffix = "YT"
            dataSource = "YT"
        Case 6
            varsuffix = "FL"
            dataSource = "FL"
        End Select
        Call setDatasourceVariables

        With configsheet.Shapes("clearProfileSelectionsButton")
            configsheet.Shapes("selectAllProfilesButton").Left = .Left
            configsheet.Shapes("selectAllProfilesButton").Top = .Top
            configsheet.Shapes("selectAllProfilesButton").Width = .Width
            configsheet.Shapes("selectAllProfilesButton").Height = .Height
        End With

    Next i

End Sub

Sub newLoginButtonsModulesSheet(Optional newSuffix As String = "MC", Optional oldSuffix As String = "AW")
    On Error Resume Next

    Dim newshape As Shape
    Dim sn As String
    Dim i As Integer



    For i = 1 To 12

        Select Case i
        Case 1
            sn = "bigBox"
        Case 2
            sn = "loginButton"
        Case 3
            sn = "loginButtonArrow"
        Case 4
            sn = "loginBoxNote"
        Case 5
            sn = "authStatusBox"
        Case 6
            sn = "licenseNote"
        Case 7
            sn = "logoutButton"
        Case 8
            sn = "manageLoginsButton"
        Case 9
            sn = "addLoginButton"
        Case 10
            sn = "addLoginButtonNote1"
        Case 11
            sn = "addLoginButtonNote2"
        Case 12
            sn = "buttonFC"
        Case 13
            sn = "logo"
        End Select

        With Sheets("Modules")
            .Shapes(sn & oldSuffix).Copy
            .Paste
            Set newshape = .Shapes(.Shapes.Count)

            newshape.Name = sn & newSuffix
            If newshape.OnAction <> vbNullString Then
                If Right(newshape.OnAction, 2) = oldSuffix Then
                    newshape.OnAction = Left(newshape.OnAction, Len(newshape.OnAction) - 2) & newSuffix
                End If
            End If
        End With

    Next i
End Sub
Sub t()
    Dim cBox As Shape
    With Sheets("Modules")
        For Each cBox In .Shapes
            If cBox.OnAction <> "" Then Debug.Print cBox.OnAction
            If InStr(1, cBox.OnAction, "SupermetricsDataGrabber.xlsm!") > 0 Then
                cBox.OnAction = Replace(cBox.OnAction, "SupermetricsDataGrabber.xlsm!", "")
            End If

            'If Right(cBox.Name, 2) = "GW" And Right(cBox.OnAction, 2) = "FB" Then
            '    cBox.OnAction = Replace(cBox.OnAction, "FB", "GW")
            'End If
        Next
    End With
End Sub
Sub changeShapePositionsInModulesSheet()
    On Error Resume Next
    Dim vasen As Double
    Dim vasen2 As Double
    Dim yla As Double
    Dim ylaalku As Double
    Dim vali As Double
    Dim i As Long
    Dim origValue As Double


    ylaalku = 120
    vali = 10

    Dim row As Integer
    Dim col As Integer

    row = 1
    col = 1

    For i = 1 To 11

        If col >= 3 Then
            col = 1
            row = row + 1
        Else
            col = col + 1
        End If

        Select Case i
        Case 1
            varsuffix = ""
            dataSource = "GA"
        Case 2
            varsuffix = "AW"
            dataSource = "AW"
        Case 3
            varsuffix = "FB"
            dataSource = "FB"
        Case 4
            varsuffix = "FA"
            dataSource = "FA"
        Case 5
            varsuffix = "AC"
            dataSource = "AC"
        Case 9
            varsuffix = "YT"
            dataSource = "YT"
        Case 6
            varsuffix = "TW"
            dataSource = "TW"

        Case 8
            varsuffix = "GW"
            dataSource = "GW"
        Case 11
            varsuffix = "ST"
            dataSource = "ST"
        Case 10
            varsuffix = "MC"
            dataSource = "MC"
             Case 7
            varsuffix = "TA"
            dataSource = "TA"
            '         Case 7
            '       varsuffix = "FL"
            '       dataSource = "FL"
        End Select

        Modules.Shapes("bigBox" & varsuffix).Width = Modules.Shapes("bigBox").Width

        yla = ylaalku + (row - 1) * Modules.Shapes("bigBox" & varsuffix).Height
        vasen = 20 + Modules.Shapes("bigBox" & varsuffix).Width * (col - 1)


        vasen2 = vasen + 15
        Dim bWidth As Double
        Dim bHeight As Double



        With Modules
            bWidth = .Shapes("logoutButton").Width
            bHeight = .Shapes("logoutButton").Height

            .Shapes("logoutButton" & varsuffix).Width = bWidth
            .Shapes("logoutButton" & varsuffix).Height = bHeight
            .Shapes("manageLoginsButton" & varsuffix).Width = bWidth
            .Shapes("manageLoginsButton" & varsuffix).Height = bHeight
            .Shapes("addLoginButton" & varsuffix).Width = bWidth
            .Shapes("addLoginButton" & varsuffix).Height = bHeight


            .Shapes("bigBox" & varsuffix).Top = yla
            .Shapes("loginButton" & varsuffix).Top = yla + 50
            .Shapes("loginButtonArrow" & varsuffix).Top = .Shapes("loginButton" & varsuffix).Top + .Shapes("loginButton" & varsuffix).Height / 2 - .Shapes("loginButtonArrow" & varsuffix).Height / 2
            .Shapes("loginBoxNote" & varsuffix).Top = .Shapes("loginButton" & varsuffix).Top + .Shapes("loginButton" & varsuffix).Height / 2 - .Shapes("loginBoxNote" & varsuffix).Height / 2
            .Shapes("authStatusBox" & varsuffix).Top = yla + 35
            .Shapes("licenseNote" & varsuffix).Top = .Shapes("authStatusBox" & varsuffix).Top + 15
            .Shapes("logoutButton" & varsuffix).Top = .Shapes("licenseNote" & varsuffix).Top + 30
            '  If dataSource = "GA" Then
            .Shapes("manageLoginsButton" & varsuffix).Top = .Shapes("logoutButton" & varsuffix).Top
            .Shapes("addLoginButton" & varsuffix).Top = .Shapes("logoutButton" & varsuffix).Top
            .Shapes("addLoginButtonNote1" & varsuffix).Top = .Shapes("addLoginButton" & varsuffix).Top + .Shapes("addLoginButton" & varsuffix).Height / 2 - .Shapes("addLoginButtonNote1" & varsuffix).Height / 2 - 2
            .Shapes("addLoginButtonNote2" & varsuffix).Top = .Shapes("addLoginButton" & varsuffix).Top + .Shapes("addLoginButton" & varsuffix).Height / 2 - .Shapes("addLoginButtonNote2" & varsuffix).Height / 2 - 1
            ' End If
            With .Shapes("logo" & varsuffix)
                .Top = yla + 5
                '  origValue = .Height

            End With
            ' .Shapes("logo" & varsuffix).Left = .Shapes("bigBox" & varsuffix).Left + .Shapes("bigBox" & varsuffix).Width - .Shapes("logo" & varsuffix).Width - 8
            .Shapes("logo" & varsuffix).Left = vasen + 8

            .Shapes("bigBox" & varsuffix).Left = vasen
            .Shapes("loginButton" & varsuffix).Left = vasen2 + 10
            .Shapes("loginButtonArrow" & varsuffix).Left = .Shapes("loginButton" & varsuffix).Left + .Shapes("loginButton" & varsuffix).Width - .Shapes("loginButtonArrow" & varsuffix).Width - 2
            .Shapes("loginBoxNote" & varsuffix).Left = .Shapes("loginButtonArrow" & varsuffix).Left - .Shapes("loginBoxNote" & varsuffix).Width - 5
            .Shapes("logoutButton" & varsuffix).Left = vasen2
            .Shapes("authStatusBox" & varsuffix).Left = vasen2
            .Shapes("licenseNote" & varsuffix).Left = vasen2

            .Shapes("buttonFC" & varsuffix).Width = .Shapes("loginButton" & varsuffix).Width
            .Shapes("buttonFC" & varsuffix).Height = .Shapes("loginButton" & varsuffix).Height
            .Shapes("buttonFC" & varsuffix).Left = .Shapes("loginButton" & varsuffix).Left
            .Shapes("buttonFC" & varsuffix).Top = .Shapes("loginButton" & varsuffix).Top

            '  If dataSource = "GA" Then
            .Shapes("manageLoginsButton" & varsuffix).Left = vasen2 + .Shapes("logoutButton" & varsuffix).Width + 10
            .Shapes("addLoginButton" & varsuffix).Left = vasen2 + .Shapes("logoutButton" & varsuffix).Width + 10
            .Shapes("addLoginButtonNote1" & varsuffix).Left = vasen2 + .Shapes("logoutButton" & varsuffix).Width + 10 + 2
            .Shapes("addLoginButtonNote2" & varsuffix).Left = vasen2 + .Shapes("logoutButton" & varsuffix).Width + 10 + 22
            '   End If
        End With
    Next i

    Call hideFLfromModulesSheet
End Sub

Sub alignCheckboxes()
    Dim s As CheckBox
    Dim i As Integer
    i = 0
    For Each s In ActiveSheet.CheckBoxes
        i = i + 1
        s.Left = ActiveSheet.CheckBoxes("timeCB").Left
        s.Top = ActiveSheet.CheckBoxes("timeCB").Top + (i - 1) * 25
    Next

End Sub

Sub makeActionButtonsSameSize()
    Dim i As Integer
    Dim sh As Shape

    For i = 1 To 8
        Select Case i
        Case 1
            Set sh = Modules.Shapes("forumButton")
        Case 2
            Set sh = Modules.Shapes("twitterButton")
        Case 3
            Set sh = Modules.Shapes("purchasingButton")
        Case 4
            Set sh = Modules.Shapes("feedbackButton")
        Case 5
            Set sh = Modules.Shapes("refreshButton")
        Case 6
            Set sh = Modules.Shapes("exportButton")
        Case 7
            Set sh = Modules.Shapes("deleteAllReportsButton")
        Case 8
            Set sh = Modules.Shapes("copyButton")
        End Select


        '
        '        With Modules.Shapes("instructionsButton")
        '            sh.Left = .Left
        '            sh.Width = .Width
        '            sh.Height = .Height
        '            sh.Top = .Top + (.Height + 5) * i
        '        End With
        '
        Const cutoff As Integer = 4
        With Modules.Shapes("instructionsButton")
            .PickUp
            sh.Apply
            sh.Width = .Width
            sh.Height = .Height
            If i <= cutoff Then
                sh.Left = .Left + (.Width + 5) * i
            Else
                sh.Left = .Left + (.Width + 5) * (i - cutoff - 1)
            End If
            If i <= cutoff Then
                sh.Top = .Top
            Else
                sh.Top = .Top + (.Height + 5)
            End If
        End With
    Next i



End Sub



Private Sub showComboBoxSizes()

    Dim cb As Shape

    For Each cb In ActiveSheet.Shapes

        If InStr(1, cb.Name, "Drop") > 0 Then Debug.Print cb.Width

    Next

    For Each cb In ActiveSheet.Shapes

        If InStr(1, cb.Name, "Drop") > 0 Then Debug.Print cb.Height

    Next

End Sub

Sub dontMoveShapesWithCellsAllConfigSh()
    Dim s As Shape
    Dim cBox As Object
    Dim i As Long
    Dim sh As Worksheet
    For i = 1 To 9
        Select Case i
        Case 1
            Set sh = Analytics
        Case 2
            Set sh = AdWords
        Case 3
            Set sh = BingAds
        Case 4
            Set sh = Facebook
        Case 5
            Set sh = YouTube
        Case 6
            Set sh = FacebookAds
        Case 7
            Set sh = Stripe
        Case 8
            Set sh = Twitter
        Case 9
            Set sh = Webmaster
        End Select
        With sh

            For Each s In .Shapes
                If Left(s.Name, 3) = "_CB" Then
                    s.Placement = xlMoveAndSize
                Else
                    s.Placement = xlFreeFloating
                End If
            Next

        End With
    Next i
End Sub
Sub MoveShapesWithCellsAllConfigSh()
    Dim s As Shape
    Dim cBox As Object
    Dim i As Long
    Dim sh As Worksheet
    For i = 1 To 5
        Select Case i
        Case 1
            Set sh = Analytics
        Case 2
            Set sh = AdWords
        Case 3
            Set sh = BingAds
        Case 4
            Set sh = Facebook
        Case 5
            Set sh = YouTube
        End Select
        With sh

            For Each s In .Shapes
                s.Placement = xlMove
            Next

        End With
    Next i
End Sub

Sub showAllConfigSheets()

    Dim i As Integer
    Dim sh As Worksheet
    For i = 1 To 5
        Select Case i
        Case 1
            Set sh = Analytics
        Case 2
            Set sh = AdWords
        Case 3
            Set sh = BingAds
        Case 4
            Set sh = Facebook
        Case 5
            Set sh = YouTube
        End Select
        With sh

            .Visible = xlSheetVisible

        End With
    Next i
End Sub
Sub dontMoveShapesWithCells()

    Dim s As Shape
    Dim cBox As Object

    With ActiveSheet

        For Each s In .Shapes
            s.Placement = xlFreeFloating
        Next

    End With
End Sub
Sub MoveShapesWithCells()

    Dim s As Shape
    Dim cBox As Object

    With ActiveSheet

        For Each s In .Shapes
            s.Placement = xlMove
        Next

    End With
End Sub
Sub moveFieldSelectionShapes()
    Dim moveup As Integer
    Dim moveleft As Integer
    Dim i As Integer
    Dim j As Integer
    Dim sn As String
    moveup = 3
    moveleft = 0
    On Error Resume Next
    With ActiveSheet
        For i = 1 To 9
            Select Case i
            Case 1
                sn = "dimbox"
            Case 2
                sn = "sdimbox"
                sn = "x"
            Case 3
                sn = "metbox"
            Case 4
                sn = "sdCategoriesNote"
            Case 5
                sn = "fieldHelpButton"
            Case 6
                sn = "drsdfb_"
                sn = "x"
            Case 7
                sn = "drdyt_"
            Case 8
                sn = "drmyt_"
            Case 9
                sn = "sdCategoriesDropdown"
            End Select
            If i < 6 Then
                .Shapes(sn).Top = .Shapes(sn).Top - moveup
                .Shapes(sn).Left = .Shapes(sn).Left - moveleft
            Else
                .DropDowns(sn).Top = .DropDowns(sn).Top - moveup
                .DropDowns(sn).Left = .DropDowns(sn).Left - moveleft
                For j = 1 To 20
                    .DropDowns(sn & j).Top = .DropDowns(sn & j).Top - moveup
                    .DropDowns(sn & j).Left = .DropDowns(sn & j).Left - moveleft
                Next j
            End If
        Next i
    End With
End Sub
Sub changeFontsForAllShapes()
    On Error Resume Next
    Dim cBox As Object
    With ActiveSheet
        For Each cBox In .Shapes
            cBox.TextFrame.Characters.Font.Name = fontName
        Next
    End With
End Sub
Sub copyShapePositionsAndFormats()
    Dim i As Integer
    Dim sFound As Boolean
    Dim s As Object
    Dim ss As Object
    Dim d As Object
    Dim sd As Object
    Dim copyFromSheet As Worksheet
    Call unprotectSheets
    Set copyFromSheet = Webmaster

    On Error Resume Next
    For i = 7 To 7

        Select Case i
        Case 1
            varsuffix = ""
            dataSource = "GA"


        Case 2
            varsuffix = "AW"
            dataSource = "AW"
        Case 3
            varsuffix = "FB"
            dataSource = "FB"
        Case 4
            varsuffix = "AC"
            dataSource = "AC"
        Case 5
            varsuffix = "YT"
            dataSource = "YT"
        Case 6
            varsuffix = "FL"
            dataSource = "FL"
        Case 7
            varsuffix = "ST"
            dataSource = "ST"
        End Select
        Call setDatasourceVariables



        With configsheet
            .Visible = xlSheetVisible
            .Rows.RowHeight = 15
            'Analytics.Cells.Copy
            '.Cells.PasteSpecial xlFormats

            For Each s In .Shapes
                sFound = False
                If shapeExists(s.Name, copyFromSheet.Name) Then
                    Set ss = copyFromSheet.Shapes(s.Name)
                    sFound = True
                ElseIf Right(s.Name, 2) = varsuffix And shapeExists(Left(s.Name, Len(s.Name) - 2), copyFromSheet.Name) Then
                    Set ss = copyFromSheet.Shapes(Left(s.Name, Len(s.Name) - 2))
                    sFound = True

                End If

                If sFound Then
                    ss.PickUp
                    s.Apply
                    s.AutoShapeType = ss.AutoShapeType
                    '    If Abs(s.Top - ss.Top) < 20 Then s.Top = ss.Top
                    s.Top = ss.Top
                    s.Left = ss.Left
                    s.Width = ss.Width
                    s.Height = ss.Height
                    '                    If s.TextFrame.Characters.Text <> ss.TextFrame.Characters.Text And UCase(s.TextFrame.Characters.Text) = ss.TextFrame.Characters.Text Then
                    '                        s.TextFrame.Characters.Text = UCase(s.TextFrame.Characters.Text)
                    '                    ElseIf ss.TextFrame.Characters.Text = "" Then
                    '                        s.TextFrame.Characters.Text = ""
                    '                    End If
                End If
            Next

            For Each s In .DropDowns
                sFound = False
                If shapeExists(s.Name, copyFromSheet.Name) Then
                    Set ss = copyFromSheet.Shapes(s.Name)
                    sFound = True
                ElseIf LCase(Right(s.Name, 2)) = LCase(varsuffix) And shapeExists(Left(s.Name, Len(s.Name) - 2), copyFromSheet.Name) Then
                    Set ss = copyFromSheet.Shapes(Left(s.Name, Len(s.Name) - 2))
                    sFound = True
                ElseIf LCase(Right(s.Name, 2)) = LCase(varsuffix) And shapeExists(Left(s.Name, Len(s.Name) - 2) & "ga", copyFromSheet.Name) Then
                    Set ss = copyFromSheet.Shapes(Left(s.Name, Len(s.Name) - 2) & "ga")
                    sFound = True
                ElseIf InStr(1, LCase(s.Name), LCase(varsuffix)) > 0 And shapeExists(Replace(s.Name, LCase(varsuffix), "ga"), copyFromSheet.Name) Then
                    Set ss = copyFromSheet.Shapes(Replace(s.Name, LCase(varsuffix), "ga"))
                    sFound = True
                End If

                If sFound Then
                    '  s.Top = ss.Top
                    s.Left = ss.Left
                    s.Width = ss.Width
                    s.Height = ss.Height

                End If
            Next
            '
            '
            '            copyFromSheet.Cells.Copy
            '            .Cells.PasteSpecial Paste:=xlFormats
            '            .Cells.PasteSpecial Paste:=xlValue
            copyFromSheet.Columns("A:Z").Copy
            .Columns("A:Z").PasteSpecial Paste:=xlPasteColumnWidths
        End With
        '
        '        For Each ss In copyFromSheet.Shapes
        '            If Left(ss.Name, 3) <> "_CB" And Not (shapeExists(ss.Name, configsheet.Name) Or shapeExists(ss.Name & varsuffix, configsheet.Name) Or shapeExists(Replace(ss.Name, "ga", LCase(varsuffix)), configsheet.Name)) Then
        '                ss.Copy
        '                configsheet.Paste
        '                Set s = configsheet.Shapes(configsheet.Shapes.Count)
        '                s.Top = ss.Top
        '                s.Left = ss.Left
        '            End If
        '        Next
        '
        With configsheet
            .Shapes("createdby").Delete
            .Shapes("selectFieldsBox").Delete
            .Shapes("actionsBox").Delete
            .Shapes("filterBox").Delete
            .Shapes("reportOptionsLabel").Delete
            .Shapes("filterButtonLogo").Delete


        End With


    Next i

End Sub

Sub updateLinkedMacros()
    Dim cBox As Object
    Dim oldSuffix As String
    Dim i As Integer
    oldSuffix = "GA"
    For i = 1 To 5

        Select Case i
        Case 1
            varsuffix = ""
            dataSource = "GA"
        Case 2
            varsuffix = "AW"
            dataSource = "AW"
        Case 3
            varsuffix = "FB"
            dataSource = "FB"
        Case 4
            varsuffix = "AC"
            dataSource = "AC"
        Case 5
            varsuffix = "YT"
            dataSource = "YT"
        Case 6
            varsuffix = "FL"
            dataSource = "FL"
        End Select
        Call setDatasourceVariables

        'Call deleteProfileSelectionCBs

        With configsheet
            .Visible = xlSheetVisible
            .Select
            .Columns("G").Font.ColorIndex = 2
            .Range("G12:G100").value = " "
            For Each cBox In .Shapes
                '                If Right(cBox.OnAction, 2) = oldSuffix Then
                '                    cBox.OnAction = Left(cBox.OnAction, Len(cBox.OnAction) - 2) & varsuffix
                '                End If
                If cBox.Fill.ForeColor.RGB = RGB(0, 101, 166) Then
                    cBox.Fill.ForeColor.RGB = RGB(242, 242, 242)
                    cBox.TextFrame.Characters.Font.ColorIndex = 1

                End If
                If cBox.AutoShapeType = 5 Then
                    cBox.Line.Visible = False
                End If
            Next

            With .DropDowns("compareDatesDD")
                configsheet.DropDowns("comparisonValueTypeDD").Width = .Width
                configsheet.DropDowns("comparisonValueTypeDD").Height = .Height
                configsheet.DropDowns("comparisonValueTypeDD").Left = .Left
            End With

            With .DropDowns("condFormDropDown")
                configsheet.DropDowns("groupingDD").Width = .Width
                configsheet.DropDowns("groupingDD").Height = .Height
                configsheet.DropDowns("groupingDD").Left = .Left
            End With
        End With
    Next i
End Sub

Sub moveAllShapes()

    Dim s As Shape
    Dim cBox As Object
    Dim movement As Double
    movement = 0

    With ActiveSheet

        For Each s In .Shapes
            s.Left = s.Left + movement
            s.Top = s.Top
        Next

        For Each cBox In .DropDowns
            cBox.Left = cBox.Left + movement
            cBox.Top = cBox.Top
        Next

    End With
End Sub



Sub change_palette_color()
    Dim color_index As Long
    color_index = 19
    ActiveWorkbook.Colors(color_index) = RGB(255, 147, 147)
    Call showcolours
End Sub


Sub showcolours()
'places colours to column A

    Dim i As Integer

    For i = 1 To 56
        Cells(i, 1).Interior.ColorIndex = i
        Cells(i, 1).value = Cells(i, 1).Interior.Color
    Next i

End Sub


Sub change_palette_all_colors()

'takes colours from column B!!!

    Dim i As Integer

    For i = 1 To 56
        ThisWorkbook.Colors(i) = Cells(i, 2).Interior.Color
    Next i

    '47 = FB

End Sub


Sub fillSettingsShColours()
    With Range("seriesColoursStart")
        .Interior.Color = RGB(230, 0, 0)
        .Offset(1).Interior.Color = RGB(0, 112, 192)
        .Offset(2).Interior.Color = RGB(122, 188, 50)
        .Offset(3).Interior.Color = RGB(255, 192, 0)
        .Offset(4).Interior.Color = RGB(112, 48, 160)
        .Offset(5).Interior.Color = RGB(13, 13, 13)
        .Offset(6).Interior.Color = RGB(146, 208, 80)
        .Offset(7).Interior.Color = RGB(151, 72, 7)
        .Offset(8).Interior.Color = RGB(85, 142, 213)
        .Offset(9).Interior.Color = RGB(245, 117, 11)
        .Offset(10).Interior.Color = RGB(255, 51, 51)
        .Offset(11).Interior.Color = RGB(127, 127, 127)
    End With
End Sub

Sub showRangeNames()

    Sheets("vars").Select
    On Error Resume Next
    Dim rivi As Long
    For rivi = 1 To vikarivi(Cells(1, 2))

        Cells(rivi, 1).value = Cells(rivi, 2).Name.value

    Next rivi

End Sub

Sub cleanAllSheets()
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        sh.Visible = xlSheetVisible
        sh.Select
        Call DeleteDirty
    Next sh
End Sub

Sub DeleteDirty()
    Dim lLastRow As Long
    Dim lLastColumn As Long
    Dim lRealLastRow As Long
    Dim lRealLastColumn As Long

    If Workbooks.Count = 0 Then
        MsgBox "No Workbooks are Currently Open!", vbInformation
        Exit Sub
    End If
    ActiveSheet.Unprotect
    If ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell) = ActiveSheet.Range("A1") Then
        Exit Sub
    End If

    With Range("A1").SpecialCells(xlCellTypeLastCell)
        lLastRow = .row
        lLastColumn = .Column
    End With

    lRealLastRow = Cells.Find("*", Range("a1"), xlFormulas, , xlByRows, xlPrevious).row
    lRealLastColumn = Cells.Find("*", Range("a1"), xlFormulas, , xlByColumns, xlPrevious).Column
    If lRealLastRow < lLastRow Then
        Range(Cells(lRealLastRow + 1, 1), Cells(lLastRow, 1)).EntireRow.Delete
    End If
    If lRealLastColumn < lLastColumn Then
        Range(Cells(1, lRealLastColumn + 1), Cells(1, lLastColumn)).EntireColumn.Delete
    End If
    ActiveSheet.UsedRange    'Resets LastCell

End Sub




Sub copyEveythingToNewSheet()

    Dim sh As Worksheet
    Dim sha As Shape
    Set sh = ActiveSheet

    With AdWords
        .Select
        Call showAllShapes
        sh.Select
        .Cells.Copy
        Cells(1, 1).Select
        sh.Paste
        For Each sha In sh.Shapes
            sha.Delete
        Next

        For Each sha In .Shapes
            sha.Copy
            sh.Paste
            With sh.Shapes(sh.Shapes.Count)
                .Top = sha.Top
                .Left = sha.Left
            End With

            Dim n As Name
            Dim nname As String
            Dim ref As String
            For Each n In ThisWorkbook.Names
                If n.RefersToRange.Worksheet.Name = "AdWords" Then
                    nname = n.Name
                    ref = n.RefersToRange.Address
                    n.Delete
                    ThisWorkbook.Names.Add nname, sh.Range(ref)
                End If
            Next n


        Next

    End With


End Sub


Sub copyEverythingToNewWB()


    Application.ScreenUpdating = False
    On Error GoTo errhandler
    Dim newWB As Workbook
    Set newWB = Workbooks.Add
    Dim newSheet As Worksheet
    Dim she As Worksheet
    Dim sh As Shape
    Dim sh2 As Shape
    Dim showGridlines
    Dim showRuler
    Dim showHeadings
    Dim i As Long
    Dim shname As String
    For Each she In ThisWorkbook.Sheets
        Set newSheet = newWB.Sheets.Add
        newSheet.Name = she.Name
        she.Cells.Copy
        With newSheet.Cells
            .PasteSpecial Paste:=xlPasteFormulas
            .PasteSpecial Paste:=xlPasteFormats
            .PasteSpecial Paste:=xlPasteColumnWidths
        End With

        For Each sh In she.Shapes
            If InStr(1, sh.Name, "Comment") = 0 Then
                sh.Copy
                newWB.Activate
                newSheet.Select
                ActiveSheet.Paste
                Set sh2 = newSheet.Shapes(newSheet.Shapes.Count)
                sh2.Top = sh.Top
                sh2.Left = sh.Left
            End If
        Next

        For Each sh In newSheet.Shapes
            If sh.OnAction <> vbNullString Then
                If InStr(1, sh.OnAction, "!") > 0 Then
                    sh.OnAction = Right(sh.OnAction, Len(sh.OnAction) - InStr(1, sh.OnAction, "!"))
                End If
            End If
            On Error Resume Next
            sh.ControlFormat.LinkedCell = Replace(Right(sh.ControlFormat.LinkedCell, Len(sh.ControlFormat.LinkedCell) - InStr(1, sh.ControlFormat.LinkedCell, "]")), "'", "")
            sh.ControlFormat.ListFillRange = Replace(Right(sh.ControlFormat.ListFillRange, Len(sh.ControlFormat.ListFillRange) - InStr(1, sh.ControlFormat.ListFillRange, "]")), "'", "")
        Next



        newSheet.Tab.ColorIndex = she.Tab.ColorIndex
        newSheet.Tab.Color = she.Tab.Color
        she.Activate
        showGridlines = ActiveWindow.DisplayGridlines
        '     showRuler = ActiveWindow.DisplayRuler
        showHeadings = ActiveWindow.DisplayHeadings
        newWB.Activate
        newSheet.Select
        ActiveWindow.DisplayGridlines = showGridlines
        '    ActiveWindow.DisplayRuler = showRuler
        ActiveWindow.DisplayHeadings = showHeadings
        ThisWorkbook.Activate
        ' newsheet.Visible = she.Visible

        For i = 1 To 56
            newWB.Colors(i) = ThisWorkbook.Colors(i)
        Next i


    Next

    Application.DisplayAlerts = False
    newWB.Sheets("Sheet1").Delete
    newWB.Sheets("Sheet2").Delete
    newWB.Sheets("Sheet3").Delete
    Application.DisplayAlerts = True

    For Each newSheet In newWB.Sheets
        newSheet.move after:=newWB.Sheets(ThisWorkbook.Sheets(ThisWorkbook.Sheets(newSheet.Name).Index - 1).Name)
    Next
    For Each newSheet In newWB.Sheets
        newSheet.move after:=newWB.Sheets(ThisWorkbook.Sheets(ThisWorkbook.Sheets(newSheet.Name).Index - 1).Name)
    Next
    For Each newSheet In newWB.Sheets
        newSheet.move after:=newWB.Sheets(ThisWorkbook.Sheets(ThisWorkbook.Sheets(newSheet.Name).Index - 1).Name)
    Next
    For Each newSheet In newWB.Sheets
        newSheet.move after:=newWB.Sheets(ThisWorkbook.Sheets(ThisWorkbook.Sheets(newSheet.Name).Index - 1).Name)
    Next
    For Each newSheet In newWB.Sheets
        newSheet.move after:=newWB.Sheets(ThisWorkbook.Sheets(ThisWorkbook.Sheets(newSheet.Name).Index - 1).Name)
    Next
    For Each newSheet In newWB.Sheets
        newSheet.move after:=newWB.Sheets(ThisWorkbook.Sheets(ThisWorkbook.Sheets(newSheet.Name).Index - 1).Name)
    Next


    For Each newSheet In newWB.Sheets
        For Each sh In newSheet.Shapes
            shname = ""
            shname = Left(sh.ControlFormat.LinkedCell, InStr(1, sh.ControlFormat.LinkedCell, "!"))
            If shname <> "" Then
                sh.ControlFormat.LinkedCell = Right(sh.ControlFormat.LinkedCell, Len(sh.ControlFormat.LinkedCell) - InStr(1, sh.ControlFormat.LinkedCell, "!"))
                sh.ControlFormat.ListFillRange = Right(sh.ControlFormat.ListFillRange, Len(sh.ControlFormat.ListFillRange) - InStr(1, sh.ControlFormat.ListFillRange, "!"))

                sh.ControlFormat.LinkedCell = shname & sh.ControlFormat.LinkedCell
                sh.ControlFormat.ListFillRange = shname & sh.ControlFormat.ListFillRange
            End If
        Next
    Next



    Dim n As Name
    For Each n In ThisWorkbook.Names
        newWB.Names.Add n.Name, newWB.Sheets(n.RefersToRange.Worksheet.Name).Range(n.RefersToRange.Address)
    Next n

    For Each newSheet In newWB.Sheets
        newSheet.Cells.Replace What:=ThisWorkbook.Path & "*[" & ThisWorkbook.Name & "]", Replacement:="", LookAt:=xlPart, _
                               MatchCase:=False
        newSheet.Cells.Replace What:="[" & ThisWorkbook.Name & "]", Replacement:="", LookAt:=xlPart, _
                               MatchCase:=False
    Next

    'PLEASE NOTE : You must choose Tools, References and select the
    'Microsoft Visual Basic for Applications Extensibility library


    'UNCOMMENT BELOW!!!
    'Dim modu As vbcomponent


    'Dim strFolder As String, strTempFile As String
    '    strFolder = ThisWorkbook.Path
    '    If Len(strFolder) = 0 Then strFolder = CurDir
    '    strFolder = strFolder & "\"
    '    strTempFile = strFolder & "~tmpexport.bas"
    ''    On Error Resume Next
    'For Each modu In ThisWorkbook.VBProject.vbcomponents
    '    ThisWorkbook.VBProject.vbcomponents(modu.Name).Export strTempFile
    '    newwb.VBProject.vbcomponents.Import strTempFile
    '    Kill strTempFile
    'Next

    Exit Sub

errhandler:
    Debug.Print Err.Description
    Resume Next
End Sub

