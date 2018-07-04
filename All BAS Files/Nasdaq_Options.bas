Attribute VB_Name = "Nasdaq_Options"
Option Explicit
Sub GetOptionChainNasdaq(Symbol As String)
    Dim threads As ThreadManager_Options
    Set threads = New ThreadManager_Options
    Dim tmpArr(0 To 0) As String, threadId As Long
    tmpArr(0) = "http://www.nasdaq.com/symbol/" & Symbol & "/option-chain?money=all"
    
    threadId = threads.startThread("GetRepsonseText", tmpArr)
    While threads.isThreadRunning(threadId)
        Application.Wait 1
    Wend
    
    Dim result As String
    result = threads.getResult(threadId)
    Dim htmlObject As New HTMLDocument, fsObject As New FileSystemObject
    With fsObject.OpenTextFile(result, ForReading)
        htmlObject.body.innerHTML = .ReadAll
        .Close
    End With
    
    Dim spot
    spot = GetNasdaqSpot(htmlObject)
    
    Dim dateLinks, pageLinks
    dateLinks = GetLinks(htmlObject, "dateindex=")
    pageLinks = GetLinks(htmlObject, "page=")
    Dim taskIds As New Collection, taskId, i As Long, j As Long, k As Long
    For i = 1 To UBound(pageLinks)
        tmpArr(0) = pageLinks(i)
        Debug.Print tmpArr(0)
        threadId = threads.startThread("GetRepsonseText", tmpArr)
        taskIds.Add threadId
    Next i
    Dim dict As New Dictionary
    dict.Add htmlObject, taskIds
    Set taskIds = New Collection
    For i = 1 To UBound(dateLinks)
        If InStr(dateLinks(i), "dateindex=-") = 0 Then
            tmpArr(0) = dateLinks(i)
            threadId = threads.startThread("GetRepsonseText", tmpArr)
            taskIds.Add threadId
        End If
    Next i
    While threads.getRuningThreads
        Application.Wait 1
    Wend
    Dim pageTaskIds As Collection
    For Each taskId In taskIds
        result = threads.getResult(CDbl(taskId))
        Set htmlObject = New HTMLDocument
        With fsObject.OpenTextFile(result, ForReading)
            htmlObject.body.innerHTML = .ReadAll
            .Close
        End With
        Set pageTaskIds = New Collection
        pageLinks = GetLinks(htmlObject, "page=")
        For i = 1 To UBound(pageLinks)
            tmpArr(0) = pageLinks(i)
            threadId = threads.startThread("GetRepsonseText", tmpArr)
            pageTaskIds.Add threadId
        Next i
        dict.Add htmlObject, pageTaskIds
    Next
    
    While threads.getRuningThreads
        Application.Wait 1
    Wend
    
    Dim headerArr()
    headerArr = Array("Calls", "Last", "Chg", "Bid", "Ask", "Vol", "Open Int", "Root", "Strike", "Puts", "Last", "Chg", "Bid", "Ask", "Vol", "Open Int")
    Dim optionTable, resultsCollection As New Collection, dictItem
    For Each dictItem In dict.Keys
        Set htmlObject = dictItem
        optionTable = GetTableValues(htmlObject, headerArr)
        If IsArray(optionTable) Then resultsCollection.Add optionTable
        For Each taskId In dict(dictItem)
            result = threads.getResult(CDbl(taskId))
            Set htmlObject = New HTMLDocument
            With fsObject.OpenTextFile(result, ForReading)
                htmlObject.body.innerHTML = .ReadAll
                .Close
            End With
            optionTable = GetTableValues(htmlObject, headerArr)
            If IsArray(optionTable) Then resultsCollection.Add optionTable
        Next
    Next
    
    Dim lastRow As Long, collectionItem
    For Each collectionItem In resultsCollection
        lastRow = lastRow + UBound(collectionItem)
    Next
    ReDim resultsArray(1 To lastRow, 1 To UBound(headerArr) + 1)
    For Each collectionItem In resultsCollection
        For i = 1 To UBound(collectionItem)
            k = k + 1
            For j = 1 To UBound(collectionItem, 2)
                resultsArray(k, j) = collectionItem(i, j)
            Next
        Next
    Next
    
    Range("nsymbol").Offset(0, 1) = spot
    Range(Range("nsymbol").Offset(1, 0).Offset(lastRow, 0), Range("nsymbol").Offset(1, 0).Offset(Rows.Count - 4, UBound(headerArr))).ClearContents
    Range(Range("nsymbol").Offset(2, 0), Range("nsymbol").Offset(2, 0).Offset(0, UBound(headerArr))) = headerArr
    Range(Range("nsymbol").Offset(3, 0), Range("nsymbol").Offset(3, 0).Offset(lastRow - 1, UBound(headerArr))) = resultsArray
    
    Set threads = Nothing
    Set htmlObject = Nothing
    Set fsObject = Nothing
    Set resultsCollection = Nothing
    Set taskIds = Nothing
    Set pageTaskIds = Nothing
    Set dict = Nothing
    
End Sub

Private Function GetLinks(htmlObject As HTMLDocument, LookFor As String)
    Dim htmlLinks As Object, htmlLink As HTMLAnchorElement, i As Long, j As Long
    Dim dict As New Dictionary
    ReDim tmpArr(0 To 0) As String
    Set htmlLinks = htmlObject.all.tags("a")
    If htmlLinks Is Nothing Then
        GetLinks = tmpArr
        Exit Function
    End If
    For Each htmlLink In htmlLinks
        If Not dict.Exists(htmlLink.href) Then
            i = InStr(htmlLink.href, LookFor)
            If i > 0 Then
                j = j + 1
                ReDim Preserve tmpArr(0 To j)
                tmpArr(j) = htmlLink.href
                dict.Add tmpArr(j), ""
            End If
        End If
    Next
    GetLinks = tmpArr
End Function
Private Function GetTableValues(htmlObject As HTMLDocument, tableHeader())
    Dim htmlTables As Object, tableObject As htmlTable, tableRow As HTMLTableRow, tableCell As HTMLTableCell
    Dim tableFound As Boolean, i As Long, j As Long
    Set htmlTables = htmlObject.all.tags("table")
    If htmlTables Is Nothing Then Exit Function
    For Each tableObject In htmlTables
        If tableObject.Rows.Length > 1 Then
            Set tableRow = tableObject.Rows(0)
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
            End If
        End If
        If tableFound Then
            ReDim tmpArr(1 To tableObject.Rows.Length - 1, 1 To UBound(tableHeader) + 1)
            For i = 1 To tableObject.Rows.Length - 1
                Set tableRow = tableObject.Rows(i)
                j = 0
                For Each tableCell In tableRow.Cells
                    j = j + 1
                    If j > UBound(tmpArr, 2) Then Exit For
                    tmpArr(i, j) = tableCell.innerText
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

Private Function GetNasdaqSpot(htmlObject As HTMLDocument)
    
    On Error Resume Next
    GetNasdaqSpot = htmlObject.getElementById("qwidget_lastsale").innerText
    If Len(GetNasdaqSpot) > 0 Then
        If IsNumeric(GetNasdaqSpot) Then GetNasdaqSpot = CDbl(GetNasdaqSpot)
    End If
    
End Function
