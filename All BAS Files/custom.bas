Attribute VB_Name = "custom"
Option Private Module
Option Explicit

Sub markprofiles()

    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    Dim rivi As Long
    Dim vrivi As Long
    Dim profID As Long
    Dim profIDcol As Long

    profIDcol = Range("profileliststart").Column + 2

    Application.EnableEvents = False

    With Sheets("profileselections")
        vrivi = vikarivi(.Cells(1, 1))
        .Columns("B").ClearContents
        For rivi = 1 To vrivi
            If rivi = vrivi Then Application.EnableEvents = True
            If .Cells(rivi, 1).value <> vbNullString Then
                profID = .Cells(rivi, 1).value
                If Application.CountIf(Analytics.Cells(1, profIDcol).EntireColumn, profID) > 0 Then
                    Analytics.Cells(Application.Match(profID, Analytics.Cells(1, profIDcol).EntireColumn, 0), profIDcol - 3).value = 1
                    Analytics.Cells(Application.Match(CStr(profID), Analytics.Cells(1, profIDcol).EntireColumn, 0), profIDcol - 3).value = 1
                    .Cells(rivi, 2).value = "Found"
                End If
            End If
        Next rivi
    End With

    Application.EnableEvents = True


End Sub



Public Function getRowDataRE(mediumRng As Range, Optional dataType As String = "rownumbers", Optional noGrouping As Boolean = False) As Variant
    On Error GoTo 0

    Dim rivi As Long
    Dim rivi2 As Long
    Dim mediumArr As Variant
    mediumArr = mediumRng.value
    Dim emptyFound As Boolean
    Dim resultArr As Variant
    Dim rowCountForEmail As Integer
    ReDim resultArr(1 To 1)
    rivi2 = 1
    emptyFound = False
    For rivi = 1 To UBound(mediumArr)
        If mediumArr(rivi, 1) = vbNullString And emptyFound = False Then
            If noGrouping Then
                rivi2 = rivi2 + 1
                ReDim Preserve resultArr(1 To rivi2)
            Else
                rivi2 = rivi2 + 2
                ReDim Preserve resultArr(1 To rivi2)
                resultArr(rivi2 - 2) = "totals_" & rowCountForEmail
            End If
            resultArr(rivi2 - 1) = ""
            resultArr(rivi2) = "grandtotals_" & mediumRng.row & "|" & mediumRng.row + rivi - 2
            rivi2 = rivi2 + 1
            emptyFound = True

        ElseIf emptyFound Then
            ReDim Preserve resultArr(1 To rivi2)
            resultArr(rivi2) = ""
            rivi2 = rivi2 + 1
        ElseIf rivi = 1 Then
            resultArr(rivi2) = rivi
            rivi2 = rivi2 + 1
            rowCountForEmail = rowCountForEmail + 1
        ElseIf mediumArr(rivi, 1) = mediumArr(rivi - 1, 1) Then    'same email
            ReDim Preserve resultArr(1 To rivi2)
            resultArr(rivi2) = rivi
            rivi2 = rivi2 + 1
            rowCountForEmail = rowCountForEmail + 1
        Else   'new email, add totals
            If noGrouping Then
                ReDim Preserve resultArr(1 To rivi2)
            Else
                ReDim Preserve resultArr(1 To rivi2 + 1)
                resultArr(rivi2) = "totals_" & rowCountForEmail
                rivi2 = rivi2 + 1
            End If
            resultArr(rivi2) = rivi
            rivi2 = rivi2 + 1
            rowCountForEmail = 1
        End If
    Next rivi


    If dataType = "addresses" Then
        For rivi2 = 1 To UBound(resultArr)
            If resultArr(rivi2) = "" Then
                resultArr(rivi2) = CVErr(xlErrNA)
            ElseIf IsNumeric(resultArr(rivi2)) Then
                resultArr(rivi2) = mediumRng.Cells(1, 1).Offset(resultArr(rivi2) - 1).Address
            Else
                resultArr(rivi2) = CVErr(xlErrNA)
            End If
        Next rivi2
    ElseIf dataType = "emailnames" Then
        For rivi2 = 1 To UBound(resultArr)
            If resultArr(rivi2) = "" Or rivi2 > UBound(mediumArr) Then
                resultArr(rivi2) = ""
            ElseIf IsNumeric(resultArr(rivi2)) Then
                resultArr(rivi2) = parseEmailName(CStr(mediumRng.Cells(1, 1).Offset(resultArr(rivi2) - 1).value))
            ElseIf Left(resultArr(rivi2), 6) = "totals" Then
                resultArr(rivi2) = resultArr(rivi2 - 1) & " Totals"
            ElseIf Left(resultArr(rivi2), 12) = "grandtotals2" Then
                resultArr(rivi2) = ""
            ElseIf Left(resultArr(rivi2), 11) = "grandtotals" Then
                resultArr(rivi2) = "Email Revenue Total"
            End If
        Next rivi2
    ElseIf dataType = "rowsabsolute" Then
        For rivi2 = 1 To UBound(resultArr)
            If resultArr(rivi2) = "" Or resultArr(rivi2) = 0 Then
                resultArr(rivi2) = CVErr(xlErrNA)
            ElseIf IsNumeric(resultArr(rivi2)) Then
                resultArr(rivi2) = resultArr(rivi2) + mediumRng.row - 1
            End If
        Next rivi2
    Else
        For rivi2 = 1 To UBound(resultArr)
            If resultArr(rivi2) = "" Or resultArr(rivi2) = 0 Then
                resultArr(rivi2) = CVErr(xlErrNA)
            End If
        Next rivi2
    End If


    getRowDataRE = Application.Transpose(resultArr)
End Function

Public Function parseEmailName(str As String) As String
    On Error GoTo errhandler
    Dim tempStr As String
    tempStr = Mid(str, InStr(1, str, "_") + 1)
    parseEmailName = Left(tempStr, InStr(1, tempStr, "_") - 1)
    Exit Function
errhandler:
    parseEmailName = str
End Function


Public Function getMetric(rowStr As Variant, metricColumn As Variant)
    Dim numRows As Integer
    Dim tempStr As String
    Dim sRow As Integer
    Dim eRow As Integer
    If Application.IsError(rowStr) Then
        getMetric = ""
        Exit Function
    End If
    With Application.Caller.Worksheet
        If Not IsNumeric(metricColumn) Then metricColumn = Range(metricColumn & "1").Column
        If Left(rowStr, 6) = "totals" Then
            numRows = CInt(Right(rowStr, Len(rowStr) - InStr(1, rowStr, "_")))
            getMetric = Application.Sum(.Range(Application.Caller.Address).Offset(-numRows).Resize(numRows, 1))
        ElseIf Left(rowStr, 11) = "grandtotals" Then
            tempStr = Mid(rowStr, InStr(1, rowStr, "_") + 1)
            sRow = CInt(Left(tempStr, InStr(1, tempStr, "|") - 1))
            eRow = CInt(Right(rowStr, Len(rowStr) - InStr(1, rowStr, "|")))
            getMetric = Application.Sum(.Range(.Cells(sRow, metricColumn), .Cells(eRow, metricColumn)))
        Else
            getMetric = .Cells(rowStr, metricColumn).value
        End If
    End With
End Function




