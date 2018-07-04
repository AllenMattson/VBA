Attribute VB_Name = "Yahoo_Options"
Option Explicit
#If VBA7 Then
    Public Declare PtrSafe Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
    Public Declare PtrSafe Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
    Public Declare PtrSafe Function InternetReadFileString Lib "wininet.dll" Alias "InternetReadFile" (ByVal hFile As Long, ByVal Buffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
    Public Declare PtrSafe Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByRef pBuffer As Byte, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
    Public Declare PtrSafe Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
#Else
    Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
    Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
    Public Declare Function InternetReadFileString Lib "wininet.dll" Alias "InternetReadFile" (ByVal hFile As Long, ByVal Buffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
    Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByRef pBuffer As Byte, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
    Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
#End If
Public Const IF_FROM_CACHE = &H1000000
Public Const IF_MAKE_PERSISTENT = &H2000000
Public Const IF_NO_CACHE_WRITE = &H4000000
       
Private Const BUFFER_LEN = 256

Sub GetOptionChain()

     Dim xlApp As Object, xlSht As Object, xlRng As Range, DataCollection As Collection, Data, tmpArr
     Dim Symbol As String, Expiry As String, c As Long, r As Long, m As Long, i As Long, spot As Double

     On Error GoTo ErrHdl
     Application.Cursor = xlWait
     Set xlApp = CreateObject("Excel.Application")
     Set xlSht = xlApp.Workbooks.Add.ActiveSheet
     If xlApp.Calculation <> xlCalculationManual Then xlApp.Calculation = xlCalculationManual
     Symbol = Range("Symbol").Text
     ExtractData xlSht, "URL;https://ca.finance.yahoo.com/q/os?s=" & Symbol
     Set xlRng = xlSht.Cells.Find(What:="View By Expiration:", LookIn:=xlValues, Lookat:=xlPart, _
                                SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
     If xlRng Is Nothing Then
         Err.Raise vbObjectError + 513, , Symbol & "is an invalid symbol ..."
     Else
         Set DataCollection = New Collection
         tmpArr = Split(Replace(LCase(xlRng), "view by expiration:", ""), "|")
         Set xlRng = xlSht.Cells.Find(What:="Calls", LookIn:=xlValues, Lookat:=xlPart, _
                           SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
         If Not xlRng Is Nothing Then
             Data = xlRng.CurrentRegion
             DataCollection.Add Data
         End If
         For i = 1 To UBound(tmpArr)
             tmpArr(i) = "1 " & Trim(tmpArr(i))
             If IsDate(tmpArr(i)) Then
                 Expiry = Format(CDate(tmpArr(i)), "YYYY") & "-" & Format(CDate(tmpArr(i)), "mm")
                 xlSht.Cells.Clear
                 ExtractData xlSht, "URL;https://ca.finance.yahoo.com/q/os?s=" & Symbol & "&m=" & Expiry
                 Set xlRng = xlSht.Cells.Find(What:="Calls", LookIn:=xlValues, Lookat:=xlPart, _
                                      SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
                 If Not xlRng Is Nothing Then
                     Data = xlRng.CurrentRegion
                     DataCollection.Add Data
                 End If
             End If
         Next i
         Set xlRng = xlSht.Cells.Find(What:="(" & Symbol & ")", LookIn:=xlValues, Lookat:=xlPart, _
                               SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
         If Not xlRng Is Nothing Then
            For i = 4 To 1 Step -1
                tmpArr = Split(xlRng.Offset(i, 0), " ")
                If (IsNumeric(tmpArr(0))) Then
                    spot = tmpArr(0)
                    Exit For
                End If
            Next
         End If
     End If
     For i = 1 To DataCollection.Count
         r = r + UBound(DataCollection(i)) - 3
         c = UBound(DataCollection(i), 2) - 1
         If c > m Then m = c
     Next i
     ReDim FinalData(1 To r + 1, 1 To c)
     r = 1
     For c = 1 To UBound(DataCollection(1), 2) - 1
         FinalData(r, c) = DataCollection(1)(2, c)
     Next c
     FinalData(1, 8) = "Strike"
     For i = 1 To DataCollection.Count
         For m = 3 To UBound(DataCollection(i)) - 1
             r = r + 1
             For c = 1 To UBound(DataCollection(i), 2) - 1
                 FinalData(r, c) = DataCollection(i)(m, c)
             Next c
         Next m
     Next i
     Range("Symbol").Offset(0, 1) = spot
     Range(Range("Symbol").Offset(2, 0), Range("Symbol").Offset(UBound(FinalData) + 1, UBound(FinalData, 2) - 1)) = FinalData
     Range(Range("Symbol").Offset(UBound(FinalData) + 2, 0), _
     Range("Symbol").Offset(Rows.Count - Range("Symbol").Row, UBound(FinalData, 2))).ClearContents
     Range(Range("Symbol").Offset(1, 0), Range("Symbol").Offset(1, UBound(FinalData, 2) - 1)).ClearContents
     Range("Symbol").Offset(2, 0).CurrentRegion.Replace "Down ", "-", xlPart, xlByColumns, False
     Range("Symbol").Offset(2, 0).CurrentRegion.Replace "Up ", "", xlPart, xlByColumns, False
     
ErrHdl:
     Application.Cursor = xlDefault
     If Err.Number Then MsgBox Err.Description, vbCritical, "Get Option Chain"
     On Error Resume Next
     If Not xlApp Is Nothing Then
         xlApp.CutCopyMode = False
         xlApp.DisplayAlerts = False
         xlApp.Quit
         Set xlApp = Nothing
     End If
     If Not DataCollection Is Nothing Then Set DataCollection = Nothing
     If Not xlSht Is Nothing Then Set xlSht = Nothing
     If Not xlRng Is Nothing Then Set xlRng = Nothing

End Sub
 
Private Sub ExtractData(xlSht As Object, QueryString As String)
     With xlSht.QueryTables.Add(Connection:=QueryString, Destination:=xlSht.Range("A1"))
         .WebSelectionType = xlEntirePage
         .WebFormatting = xlWebFormattingNone
         .Refresh BackgroundQuery:=False
     End With
End Sub


Sub GetOptionChain_New()

     Dim DataCollection As Collection, Data, tmpArr
     Dim Symbol As String, Expiries() As String, c As Long, r As Long, m As Long, i As Long, spot As Double

     On Error GoTo ErrHdl
     Symbol = Range("Symbol").Text
     Set DataCollection = New Collection
         
     Expiries = GetExpiries(Symbol)
     
     If UBound(Expiries) < 2 Then
        Err.Raise vbObjectError + 513, , Symbol & "is an either invalid or does not have options ..."
     Else
        For i = 1 To UBound(Expiries)
            Data = GetOptions(Expiries(i), Symbol)
            If IsArray(Data) Then DataCollection.Add Data
        Next i
     End If
     
     For i = 1 To DataCollection.Count
         r = r + UBound(DataCollection(i))
     Next i
     ReDim FinalData(1 To r + 1, 1 To 15)
     FinalData(1, 1) = "Symbol": FinalData(1, 9) = "Symbol"
     FinalData(1, 2) = "Last": FinalData(1, 10) = "Last"
     FinalData(1, 3) = "Change": FinalData(1, 11) = "Change"
     FinalData(1, 4) = "Bid": FinalData(1, 12) = "Bid"
     FinalData(1, 5) = "Ask": FinalData(1, 13) = "Ask"
     FinalData(1, 6) = "Volume": FinalData(1, 14) = "Volume"
     FinalData(1, 7) = "Open Int": FinalData(1, 15) = "Open Int"
     FinalData(1, 8) = "Strike"
     
     r = 1
     
     For i = 1 To DataCollection.Count
         For m = 1 To UBound(DataCollection(i))
             r = r + 1
             For c = 1 To 15
                 FinalData(r, c) = DataCollection(i)(m, c)
             Next c
         Next m
     Next i
     
     Data = GetPrice(Symbol)
     If IsArray(Data) Then spot = Data(1, 1)
     
     Range("Symbol").Offset(0, 1) = spot
     Range(Range("Symbol").Offset(2, 0), Range("Symbol").Offset(UBound(FinalData) + 1, UBound(FinalData, 2) - 1)) = FinalData
     Range(Range("Symbol").Offset(UBound(FinalData) + 2, 0), _
     Range("Symbol").Offset(Rows.Count - Range("Symbol").Row, UBound(FinalData, 2))).ClearContents
     Range(Range("Symbol").Offset(1, 0), Range("Symbol").Offset(1, UBound(FinalData, 2) - 1)).ClearContents
     
ErrHdl:
     
     If Err.Number Then MsgBox Err.Description, vbCritical, "Get Option Chain"
     On Error Resume Next
     If Not DataCollection Is Nothing Then Set DataCollection = Nothing
     

End Sub


Private Function GetOptions(url As String, Symbol As String)
    
    Dim htm As HTMLDocument, htmlTables As Object, xmlHTTP As Object, PutStrikes As String, CallStrikes As String, tmpDate As String, Dates As String
    Dim i As Long, j As Long, Data, k As Long, l As Long, m As Long, ii As Long, Loops As Long, SymbolLength As Long, tempArr, callArr, putArr, dateArr
    Dim counter As Long, coll As New Collection, maxLength As Long, p As Long, c As Long
        
        
    Dim tableArr, rowArr, cellArr, tmpStr As String
    On Error GoTo ErrHdl
    
    SymbolLength = Len(Symbol)
    
    tmpStr = WebGetHTML(url)
    tableArr = Split(LCase(tmpStr), "<table ")
    rowArr = Split(tableArr(2), "<tr ")
    ReDim Calls(1 To UBound(rowArr) - 1, 1 To 8)
    For i = 2 To UBound(rowArr)
        cellArr = Split(rowArr(i), "<td")
        For j = 1 To 3
            k = InStr(cellArr(j), "</")
            tmpStr = Trim(Left(cellArr(j), k - 1))
            l = InStrRev(tmpStr, ">")
            Calls(i - 1, j) = UCase(Right(tmpStr, Len(tmpStr) - l))
        Next j
        k = InStr(cellArr(6), "</")
        tmpStr = Trim(Left(cellArr(6), k - 1))
        l = InStrRev(tmpStr, ">")
        Calls(i - 1, 4) = UCase(Right(tmpStr, Len(tmpStr) - l))
        For j = 4 To 5
            k = InStr(cellArr(j), "</")
            tmpStr = Trim(Left(cellArr(j), k - 1))
            l = InStrRev(tmpStr, ">")
            Calls(i - 1, j + 1) = UCase(Right(tmpStr, Len(tmpStr) - l))
        Next j
        For j = 8 To 9
            k = InStr(cellArr(j), "</")
            tmpStr = Trim(Left(cellArr(j), k - 1))
            l = InStrRev(tmpStr, ">")
            Calls(i - 1, j - 1) = UCase(Right(tmpStr, Len(tmpStr) - l))
        Next j
    Next i
    
    rowArr = Split(tableArr(3), "<tr ")
    ReDim Puts(1 To UBound(rowArr) - 1, 1 To 8)
    For i = 2 To UBound(rowArr)
        cellArr = Split(rowArr(i), "<td")
        For j = 1 To 3
            k = InStr(cellArr(j), "</")
            tmpStr = Trim(Left(cellArr(j), k - 1))
            l = InStrRev(tmpStr, ">")
            Puts(i - 1, j) = UCase(Right(tmpStr, Len(tmpStr) - l))
        Next j
        k = InStr(cellArr(6), "</")
        tmpStr = Trim(Left(cellArr(6), k - 1))
        l = InStrRev(tmpStr, ">")
        Puts(i - 1, 4) = UCase(Right(tmpStr, Len(tmpStr) - l))
        For j = 4 To 5
            k = InStr(cellArr(j), "</")
            tmpStr = Trim(Left(cellArr(j), k - 1))
            l = InStrRev(tmpStr, ">")
            Puts(i - 1, j + 1) = UCase(Right(tmpStr, Len(tmpStr) - l))
        Next j
        For j = 8 To 9
            k = InStr(cellArr(j), "</")
            tmpStr = Trim(Left(cellArr(j), k - 1))
            l = InStrRev(tmpStr, ">")
            Puts(i - 1, j - 1) = UCase(Right(tmpStr, Len(tmpStr) - l))
        Next j
    Next i
    
    If UBound(Calls) > UBound(Puts) Then
        maxLength = UBound(Calls)
    Else
        maxLength = UBound(Puts)
    End If
    CallStrikes = "": PutStrikes = ""
    For i = 1 To maxLength
        If i <= UBound(Calls) Then
            If InStr(CallStrikes, Calls(i, 1)) = 0 Then CallStrikes = CallStrikes & "|" & Calls(i, 1)
        End If
        If i <= UBound(Puts) Then
            If InStr(PutStrikes, Puts(i, 1)) = 0 Then PutStrikes = PutStrikes & "|" & Puts(i, 1)
        End If
    Next i
    callArr = Split(CallStrikes, "|"): putArr = Split(PutStrikes, "|")
    ReDim tempArr(0 To 0)
    j = 1: k = 1: i = 0
    While j <= UBound(callArr) And k <= UBound(putArr)
        i = i + 1
        ReDim Preserve tempArr(0 To i)
        If callArr(j) = putArr(k) Then
            tempArr(i) = callArr(j)
            j = j + 1
            k = k + 1
        ElseIf CDbl(callArr(j)) < CDbl(putArr(k)) Then
            tempArr(i) = callArr(j)
            j = j + 1
        Else
            tempArr(i) = putArr(k)
            k = k + 1
        End If
    Wend
    While j <= UBound(callArr)
        i = i + 1
        ReDim Preserve tempArr(0 To i)
        tempArr(i) = callArr(j)
        j = j + 1
    Wend
    While k <= UBound(putArr)
        i = i + 1
        ReDim Preserve tempArr(0 To i)
        tempArr(i) = putArr(k)
        k = k + 1
    Wend
    If UBound(tempArr) > 0 Then
        ReDim Data(1 To UBound(tempArr), 1 To 15)
        j = 1: k = 1
        For i = 1 To UBound(tempArr)
            If j <= UBound(Calls) Then
                If tempArr(i) = Calls(j, 1) Then
                    For l = 1 To 7
                        Data(i, l) = Calls(j, l + 1)
                        If IsNumeric(Data(i, l)) Then Data(i, l) = CDbl(Data(i, l))
                    Next l
                    Data(i, 8) = tempArr(i)
                    If IsNumeric(Data(i, 8)) Then Data(i, 8) = CDbl(Data(i, 8))
                    j = j + 1
                End If
            End If
            If k <= UBound(Puts) Then
                If tempArr(i) = Puts(k, 1) Then
                    For l = 9 To 15
                        Data(i, l) = Puts(k, l - 7)
                        If IsNumeric(Data(i, l)) Then Data(i, l) = CDbl(Data(i, l))
                    Next l
                    Data(i, 8) = tempArr(i)
                    If IsNumeric(Data(i, 8)) Then Data(i, 8) = CDbl(Data(i, 8))
                    k = k + 1
                End If
            End If
            If Len(Data(i, 1)) = 0 Then
                Data(i, 1) = Left(Data(i, 9), SymbolLength + 6) & "C" & Right(Data(i, 9), Len(Data(i, 9)) - SymbolLength - 7)
                For m = 2 To 7
                    Data(i, m) = " - "
                Next m
            ElseIf Len(Data(i, 9)) = 0 Then
                Data(i, 9) = Left(Data(i, 1), SymbolLength + 6) & "P" & Right(Data(i, 1), Len(Data(i, 1)) - SymbolLength - 7)
                For m = 10 To 15
                    Data(i, m) = " - "
                Next m
            End If
        Next i
    End If
    
    GetOptions = Data
    
ErrHdl:
    
    If Err.Number Then
        Data = Err.Description
        GetOptions = Data
    End If
    Set htm = Nothing
    Set xmlHTTP = Nothing
    Set htmlTables = Nothing

End Function

Private Function GetExpiries(Symbol As String)
    
    Dim tmpStr As String, tmpArr, i As Long, j As Long, url As String, k As Long, Data() As String
        
    On Error GoTo ErrHdl
        
    url = "https://finance.yahoo.com/q/op?s=" & Symbol
    tmpStr = WebGetHTML(url)
    tmpArr = Split(LCase(tmpStr), "<option ")
    If UBound(tmpArr) < 1 Then Exit Function
    ReDim Data(1 To UBound(tmpArr)) As String
    For i = 1 To UBound(tmpArr)
        k = InStr(LCase(tmpArr(i)), "date=")
        j = InStr(k, LCase(tmpArr(i)), Chr(34))
        tmpStr = Mid(tmpArr(i), k + 5, j - k - 5)
        Data(i) = url & "&date=" & tmpStr
    Next i
    GetExpiries = Data
    
ErrHdl:
    
    If Err.Number Then
        ReDim Data(1 To 1)
        Data(1) = Err.Description
        GetExpiries = Data
    End If
    
End Function


Function GetPrice(Symbols) As Variant
    
    Dim htm As HTMLDocument, htmlTables As Object, xmlHTTP As Object
    Dim i As Long, j As Long, url As String, Symbol, Data, Loops As Long
        
    On Error GoTo ErrHdl
        
    url = "http://finance.yahoo.com/quotes/"
    If IsArray(Symbols) Or IsObject(Symbols) Then
        For Each Symbol In Symbols
            url = url & CStr(Symbol) & ","
        Next Symbol
        url = Left(url, Len(url) - 1)
    Else
        url = url & CStr(Symbols)
    End If
    
    Set xmlHTTP = CreateObject("MSXML2.XMLHTTP")
    Set htm = New HTMLDocument
    If Not xmlHTTP Is Nothing Then
        Do
            With xmlHTTP
                .Open "GET", url, False
                .send
                htm.body.innerHTML = .responseText
            End With
            Set htmlTables = htm.all.tags("TABLE")
            If Not htmlTables Is Nothing Then
                If htmlTables.Length > 2 Then
                    If htmlTables(1).Rows.Length > 1 Then
                        If htmlTables(1).Rows(1).Cells.Length > 2 Then
                            If Len(Trim(htmlTables(1).Rows(1).Cells(2).innerText)) > 1 Then Exit Do
                        End If
                    End If
                End If
            End If
            Loops = Loops + 1
            If Loops > 10 Then Exit Do
        Loop
        If Not htmlTables Is Nothing Then
            If Not htmlTables(1) Is Nothing Then
                With htmlTables(1)
                    ReDim Data(1 To .Rows.Length - 1, 1 To .Rows(0).Cells.Length - 1)
                    For i = 1 To UBound(Data)
                        For j = 2 To UBound(Data, 2)
                            Data(i, j - 1) = .Rows(i).Cells(j).innerText
                            If IsNumeric(Data(i, j - 1)) Then Data(i, j - 1) = CDbl(Data(i, j - 1))
                        Next j
                        Data(i, UBound(Data, 2)) = .Rows(i).Cells(1).innerText
                    Next i
                End With
            End If
        End If
    End If
    GetPrice = Data
    
ErrHdl:
    
    If Err.Number Then GetPrice = "#" & Err.Description
    Set htm = Nothing
    Set xmlHTTP = Nothing
    Set htmlTables = Nothing
    
End Function

Public Function WebGetHTML(sURL As String) As String
    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
    Dim hInternet As Long, hSession As Long, lReturn As Long
    
    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
    
    If hInternet Then
        DoEvents
        iResult = InternetReadFileString(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sData = sBuffer
        
        Do While lReturn <> 0
            DoEvents
            iResult = InternetReadFileString(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sData = sData + Mid(sBuffer, 1, lReturn)
        Loop
    End If
   
    iResult = InternetCloseHandle(hInternet)
    WebGetHTML = sData
End Function


