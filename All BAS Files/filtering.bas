Attribute VB_Name = "filtering"
Option Private Module
Option Explicit

Sub clearFilters()
    On Error Resume Next
    Dim i As Long
    Application.ScreenUpdating = False
    Call unprotectSheets
    Call setDatasourceVariables
    If dataSource = "GA" Then
        Range("predeffilternum").value = 1
        With filterUF.segmentLB
            For i = 0 To .ListCount - 1
                .Selected(i) = False
            Next i
        End With
        Range("lastSegmentSelections").value = vbNullString
        Range("segmentSelectionCodes").value = vbNullString
        Range("segmentSelectionNames").value = vbNullString
    End If
    configsheet.Shapes("clearFiltersButton").Visible = False
    Range("filterstring" & varsuffix).value = vbNullString
    Range("filterRange" & varsuffix).Resize(6, 4).ClearContents

    filterStr = vbNullString
    
    configsheet.Shapes("filterNote" & varsuffix).TextFrame.Characters.Text = "No filters have been set up"
        
    Call protectSheets
End Sub

Sub launchFilterUF()
    On Error Resume Next
    If debugMode Then On Error GoTo 0
    Dim tempStr As String
    Application.ScreenUpdating = False
    Call unprotectSheets

    Call setDatasourceVariables
    Load filterUF
    Call constructListOfFilterFields

    With Sheets("vars").Range("filterRange" & varsuffix)
        For filternum = 1 To 5
            If .Offset(filternum - 1).value <> vbNullString Then
                filterUF.Controls("fieldDD" & filternum).Text = .Offset(filternum - 1).value
                filterUF.Controls("operatorDD" & filternum).Text = .Offset(filternum - 1, 1).value
                tempStr = .Offset(filternum - 1, 2).value

                '  tempStr = Replace(tempStr, "\\", "\")
                tempStr = Replace(tempStr, "\,", ",")
                tempStr = Replace(tempStr, "\;", ";")


                filterUF.Controls("value" & filternum).Text = .Offset(filternum - 1, 2).value
                Call makeFilterVisible2
                If .Offset(filternum - 1, 3).value <> "" Then

                    filterUF.Controls("and" & filternum).Visible = True
                    filterUF.Controls("or" & filternum).Visible = True
                    If .Offset(filternum - 1, 3).value = "AND" Then
                        filterUF.Controls("and" & filternum).value = True
                    Else
                        filterUF.Controls("or" & filternum).value = True
                    End If

                End If
            End If
        Next filternum
    End With

    If dataSource = "GA" Then
        Dim lastSegmentSelections As String
        Dim tempArr As Variant
        Dim i As Long
        filterUF.mpage.segmentPage.Visible = True

        With filterUF.segmentLB
            For i = 0 To .ListCount - 1
                .Selected(i) = False
                Debug.Print "SEL: " & i & "  " & .Selected(i)
            Next i

            Debug.Print "SEL: " & lastSegmentSelections

            lastSegmentSelections = Range("lastSegmentSelections").value
            If lastSegmentSelections <> vbNullString Then
                tempArr = Split(lastSegmentSelections, ",")
                For i = 0 To UBound(tempArr)
                    .Selected(tempArr(i)) = True
                    Debug.Print "SEL: " & tempArr(i) & "  " & lastSegmentSelections
                Next i
            End If
        End With
    Else
        filterUF.mpage.segmentPage.Visible = False
        filterUF.mpage.value = 0
    End If

    If dataSource = "GA" Then Application.Wait (Now + TimeValue("00:00:01"))  'prevents Excel from registering accidental 2nd click that selects segment
    filterUF.Show
    End
    Call protectSheets
End Sub
Sub makeFilterVisible2()
    With filterUF
        .Controls("operatorDD" & filternum).Visible = True
        .Controls("fieldDD" & filternum).Visible = True
        .Controls("value" & filternum).Visible = True
        If filternum < 5 Then
            If .Controls("fieldDD" & filternum + 1).Visible = False Then
                .Controls("addFilter" & filternum + 1).Visible = True
            Else
                .Controls("and" & filternum).Visible = True
                .Controls("or" & filternum).Visible = True
            End If
        End If
        If filternum > 1 Then .Controls("addFilter" & filternum).Visible = False
        .Controls("clearB" & filternum).Visible = True
    End With
End Sub
Sub constructListOfFilterFields()
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim rivi As Long
    Dim vrivi As Long
    Dim resultStr As String
    Dim filternum As Long
    Dim i As Integer
    Dim varArr() As Variant
    Dim col As Long

    Call setDatasourceVariables

    ReDim filterUFArr(1 To 1000, 1 To 4)
    '1 filter name disp
    '2 filter code
    '3 dim / met

    For i = 1 To 5
        filterUF.Controls("fieldDD" & i).Clear
    Next i
    filternum = 0



    ' If dataSource <> "YT" Then
    col = Range("dimensionsAllStart" & varsuffix).Column
    vrivi = vikarivi(Range("dimensionsAllStart" & varsuffix))
    With varsSheetForDataSource
        varArr = .Range(.Range("dimensionsAllStart" & varsuffix).Offset(2, -1), .Cells(vrivi, col + 2)).value
    End With
    '    Else
    '        Range("metricsCalcYT").Calculate
    '        col = Range("filterFieldsStartYT").Column
    '        vrivi = vikarivi(Range("filterFieldsStartYT"))
    '        With varsSheetForDataSource
    '            varArr = .Range(.Range("filterFieldsStartYT"), .Cells(vrivi, col + 2)).value
    '        End With
    '
    '    End If

    For rivi = 1 To UBound(varArr)
        filternum = filternum + 1
        For i = 1 To 5
            filterUF.Controls("fieldDD" & i).AddItem (varArr(rivi, 1))
        Next i
        filterUFArr(filternum, 1) = varArr(rivi, 1)
        filterUFArr(filternum, 2) = varArr(rivi, 2)
        filterUFArr(filternum, 3) = "dim"
    Next rivi


    '   If dataSource = "YT" Then Exit Sub

    col = Range("metricsAllStart" & varsuffix).Column
    vrivi = vikarivi(Range("metricsAllStart" & varsuffix))
    With varsSheetForDataSource
        varArr = .Range(.Range("metricsAllStart" & varsuffix).Offset(2, -1), .Cells(vrivi, col + 8)).value
    End With

    For i = 1 To 5
        filterUF.Controls("fieldDD" & i).AddItem ("")
    Next i

    For rivi = 1 To UBound(varArr)
        If (varArr(rivi, 4) = 1 Or varArr(rivi, 4) = "") And (dataSource <> "GA" Or varArr(rivi, 9) = "") Then
            filternum = filternum + 1
            For i = 1 To 5
                filterUF.Controls("fieldDD" & i).AddItem (varArr(rivi, 1))
            Next i
            filterUFArr(filternum, 1) = varArr(rivi, 1)
            filterUFArr(filternum, 2) = varArr(rivi, 2)
            filterUFArr(filternum, 3) = "met"
        End If
    Next rivi

    ' segmentLB.List = Range("segments").value

    If dataSource = "GA" Then
        varArr = Range("segments").value
        With filterUF.Controls("segmentLB")
            .Clear
            For rivi = 1 To UBound(varArr)
                .AddItem (varArr(rivi, 1))
            Next rivi
        End With
    End If


End Sub


