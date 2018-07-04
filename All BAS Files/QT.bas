Attribute VB_Name = "QT"
Option Private Module
Option Explicit
Public queryTableResultStr As String
Public qtTempSheet As Worksheet

Sub deleteQueryTables()
    Dim qt As QueryTable
    With ThisWorkbook.Sheets("QT")
        For Each qt In .QueryTables
            qt.Delete
        Next
    End With
    Call deleteDataConnections
End Sub

Sub addQueryTable(URL As String)
    Call deleteQueryTables
    If SheetExists("QT") = False Then
        Set qtTempSheet = ThisWorkbook.Sheets.Add
        qtTempSheet.Name = "QT"
        qtTempSheet.Visible = xlSheetVeryHidden
    Else
        Set qtTempSheet = ThisWorkbook.Sheets("QT")
    End If
    With qtTempSheet
        With .QueryTables.Add("URL;" & URL, .Range("A1"))
            .BackgroundQuery = False
        End With
    End With
    Range("lastQTurl").value = URL
    ' If debugMode = True Then Debug.Print "NEWQT: " & URL
End Sub

Sub fetchDataWithQueryTableDirect(URL As String, payload As String, Optional URLdecode As Boolean = False, Optional UTF8decode As Boolean = True)


    On Error Resume Next
    If debugMode = True Then On Error GoTo 0

    If URL <> Range("lastQTurl") Then Call addQueryTable(URL)

    Dim totalRows As Long
    Dim rivi As Long
    Dim sar As Long

    Dim i As Long

    Dim lastCell As Range
    Dim solu As Range

    Dim dataValues As Variant
    Dim tempStr As String

    If qtTempSheet Is Nothing Then Set qtTempSheet = ThisWorkbook.Sheets("QT")

    With qtTempSheet

        .Cells.ClearContents
        .Cells.ClearFormats

        On Error Resume Next
        With .QueryTables(1)
            .PostText = payload
            .refresh
        End With
        DoEvents


        If debugMode = True Then On Error GoTo 0

        Set lastCell = findLastCell(qtTempSheet)
        queryTableResultStr = vbNullString


        If lastCell.row = 1 And lastCell.Column = 1 Then
            queryTableResultStr = .Cells(1, 1).value
        Else
            i = 0
            dataValues = .Range(.Cells(1, 1), .Cells(lastCell.row, lastCell.Column)).value
            totalRows = UBound(dataValues, 1)
            If totalRows > 200 Then Call updateProgressAdditionalMessage("Rows to process: " & totalRows)
            For rivi = 1 To UBound(dataValues, 1)
                i = i + 1
                If i = 1000 Then
                    queryTableResultStr = queryTableResultStr & tempStr
                    tempStr = vbNullString
                    Call updateProgressAdditionalMessage("Rows processed: " & rivi & "/" & totalRows)
                    i = 0
                End If
                For sar = 1 To UBound(dataValues, 2)
                    If dataValues(rivi, sar) <> vbNullString Then tempStr = tempStr & dataValues(rivi, sar)
                Next sar
            Next rivi
            If IsArray(dataValues) Then Erase dataValues
            queryTableResultStr = queryTableResultStr & tempStr
            tempStr = vbNullString
        End If


        If queryTableResultStr = "" And Left(URL, 5) = "https" Then
            URL = "http" & Right(URL, Len(URL) - 5)
            Call fetchDataWithQueryTableDirect(URL, payload, URLdecode, UTF8decode)
            Exit Sub
        End If

        .Cells.ClearContents
        .Cells.ClearFormats

        If queryTableResultStr <> vbNullString Then

            queryTableResultStr = Replace(queryTableResultStr, vbCrLf, "")
            queryTableResultStr = Replace(queryTableResultStr, vbCr, "")
            queryTableResultStr = Replace(queryTableResultStr, vbLf, "")
            queryTableResultStr = Replace(queryTableResultStr, Chr$(9), "")
            queryTableResultStr = Replace(queryTableResultStr, Chr$(10), "")
            queryTableResultStr = Replace(queryTableResultStr, Chr$(11), "")
            queryTableResultStr = Replace(queryTableResultStr, Chr$(12), "")
            queryTableResultStr = Replace(queryTableResultStr, Chr$(13), "")
            queryTableResultStr = Replace(queryTableResultStr, Chr$(14), "")
            queryTableResultStr = Replace(queryTableResultStr, Chr$(15), "")
            queryTableResultStr = Replace(queryTableResultStr, Chr$(8), "")
            queryTableResultStr = Replace(queryTableResultStr, Chr$(127), "")

            queryTableResultStr = chrDecode(queryTableResultStr)
            If URLdecode Then queryTableResultStr = URL_decode(queryTableResultStr)
            If UTF8decode Then queryTableResultStr = UTF8_Decode(queryTableResultStr)    '

            queryTableResultStr = Trim(queryTableResultStr)

        End If

    End With


    '  If debugMode = True Then Debug.Print "Q result: " & Left$(queryTableResultStr, 1000)
End Sub
Sub deleteDataConnections()
    On Error Resume Next
    Dim objConn As Object
    For Each objConn In ActiveWorkbook.Connections
        objConn.Delete
    Next objConn
End Sub

