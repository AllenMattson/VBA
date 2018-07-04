Attribute VB_Name = "MarketWatch_Options"
Option Explicit

Sub GetOptionChainMW(Symbol As String)
    
    On Error GoTo ErrHdl
    
    Dim url As String, xmlHTTP As Object
    url = "http://www.marketwatch.com/investing/stock/" + Symbol + "/options?countrycode=US&showAll=True"
    Set xmlHTTP = CreateObject("MSXML2.XMLHTTP")
    If xmlHTTP Is Nothing Then Err.Raise vbObjectError + 513
    
    Dim htmlObject As New HTMLDocument
    With xmlHTTP
        .Open "GET", url, False
        .send
        htmlObject.body.innerHTML = .responseText
    End With
    
    Dim spot As String
    
    Dim headerArr()
    headerArr = Array("Symbol", "Last", "Change", "Vol", "Bid", "Ask", "Open Int.", "Strike", "Symbol", "Last", "Change", "Vol", "Bid", "Ask", "Open Int.")
    ReDim dataArr(1 To 1, 1 To UBound(headerArr) + 1)
    Dim optionTable
    optionTable = GetTableValues(htmlObject, headerArr)
    headerArr(0) = "Call": headerArr(8) = "Put"
    If IsArray(optionTable) Then
        Dim i As Long, j As Long, r As Long
        ReDim dataArr(1 To UBound(optionTable), 1 To UBound(optionTable, 2))
        For i = 1 To UBound(optionTable)
            If IsDate(optionTable(i, 1)) Then
                r = r + 1
                For j = 1 To UBound(optionTable, 2)
                    dataArr(r, j) = optionTable(i, j)
                Next
            ElseIf Len(spot) = 0 And IsNumeric(optionTable(i, 2)) Then
                spot = optionTable(i, 2)
            End If
        Next
    End If
    
    Range("msymbol").Offset(0, 1) = spot
    Range(Range("msymbol").Offset(1, 0).Offset(UBound(dataArr), 0), Range("msymbol").Offset(1, 0).Offset(Rows.Count - 4, UBound(headerArr))).ClearContents
    Range(Range("msymbol").Offset(2, 0), Range("msymbol").Offset(2, 0).Offset(0, UBound(headerArr) - 1)) = headerArr
    Range(Range("msymbol").Offset(3, 0), Range("msymbol").Offset(3, 0).Offset(UBound(dataArr) - 1, UBound(headerArr))) = dataArr
    
ErrHdl:
    If Err.Number Then MsgBox "Could not downlod data for " + Symbol, vbCritical, "Get Data"
    Set htmlObject = Nothing
    
End Sub

Private Function GetTableValues(htmlObject As HTMLDocument, tableHeader())
    Dim htmlTables As Object, tableObject As htmlTable, tableRow As HTMLTableRow, tableCell As HTMLTableCell
    Dim tableFound As Boolean, i As Long, j As Long, startRow As Long, htmlLink As Object
    Set htmlTables = htmlObject.all.tags("table")
    If htmlTables Is Nothing Then Exit Function
    For Each tableObject In htmlTables
        If tableObject.Rows.Length > 1 Then
            For j = 0 To tableObject.Rows.Length - 1
                Set tableRow = tableObject.Rows(j)
                If tableRow.Cells.Length > UBound(tableHeader) Then
                    tableFound = True: i = -1
                    For Each tableCell In tableRow.Cells
                        i = i + 1
                        If i > UBound(tableHeader) Then Exit For
                        If InStr(Trim(LCase(tableCell.innerText)), LCase(Trim(tableHeader(i)))) <> 1 Then
                            tableFound = False
                            Exit For
                        End If
                    Next
                    If tableFound Then Exit For
                End If
             Next
        End If
        If tableFound Then
            ReDim tmpArr(1 To tableObject.Rows.Length - 1 - j, 1 To UBound(tableHeader) + 1)
            startRow = j + 1
            For i = startRow To tableObject.Rows.Length - 1
                Set tableRow = tableObject.Rows(i)
                j = 0
                For Each tableCell In tableRow.Cells
                    j = j + 1
                    If j > UBound(tmpArr, 2) Then Exit For
                    Set htmlLink = tableCell.getElementsByTagName("a").Item(0)
                    If htmlLink Is Nothing Then
                        tmpArr(i - startRow + 1, j) = tableCell.innerText
                    Else
                        tmpArr(i - startRow + 1, j) = getExpiryDate(CStr(htmlLink.href))
                    End If
                Next
            Next
            Exit For
        End If
    Next
    If tableFound Then GetTableValues = tmpArr
    Set htmlTables = Nothing
    Set tableObject = Nothing
    Set tableRow = Nothing
    Set tableCell = Nothing
End Function

Private Function getExpiryDate(link As String) As String
    If Len(link) < 12 Then Exit Function
    Dim monthValue As Integer, monthChar As String
    monthChar = Left(Right(link, 12), 1)
    If monthChar = "A" Or monthChar = "M" Then
        monthValue = 1
    ElseIf monthChar = "B" Or monthChar = "N" Then
        monthValue = 2
    ElseIf monthChar = "C" Or monthChar = "O" Then
        monthValue = 3
    ElseIf monthChar = "D" Or monthChar = "P" Then
        monthValue = 4
    ElseIf monthChar = "E" Or monthChar = "Q" Then
        monthValue = 5
    ElseIf monthChar = "F" Or monthChar = "R" Then
        monthValue = 6
    ElseIf monthChar = "G" Or monthChar = "S" Then
        monthValue = 7
    ElseIf monthChar = "H" Or monthChar = "T" Then
        monthValue = 8
    ElseIf monthChar = "I" Or monthChar = "U" Then
        monthValue = 9
    ElseIf monthChar = "J" Or monthChar = "V" Then
        monthValue = 10
    ElseIf monthChar = "K" Or monthChar = "W" Then
        monthValue = 11
    ElseIf monthChar = "L" Or monthChar = "X" Then
        monthValue = 12
    End If
    If monthValue = 0 Then Exit Function
    Dim dayString As String, yearString As String
    dayString = Left(Right(link, 11), 2)
    If Not IsNumeric(dayString) Then Exit Function
    yearString = Left(Right(link, 9), 2)
    If Not IsNumeric(yearString) Then Exit Function
    getExpiryDate = Format(CDate(monthValue & "/" & dayString & "/20" & yearString), "dd-MMM-yy")
End Function

