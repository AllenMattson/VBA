Attribute VB_Name = "Google_Options"
Option Explicit

Sub GetOptionChain()
    
    Dim JSONtext As String, ExpirationsJSON As String, Symbol As String, spot, counter As Long
    Dim xmlHttpRequest As Object, xmlHttpReturnText As String, OptionCollection As Collection
    Dim lib As JsonLib_Options, Json As Object, JsonOpts As Object, temp, url As String
    Dim i As Long, j As Long, l As Long, jObj As Object, k, lObj As Object, Expirations As String
    Dim Calls As String, Puts As String, Data, FieldName, Keys, Caller As Range
    
    Application.EnableCancelKey = xlDisabled
    Symbol = Range("gsymbol").Text
    url = "http://www.google.com/finance/option_chain?q=" & Symbol & "&output=json"
    Set xmlHttpRequest = CreateObject("MSXML2.XMLHTTP")
    With xmlHttpRequest
        .Open "GET", url, False
        .send
        xmlHttpReturnText = .responseText
    End With
    Expirations = Replace(Trim(ExtractText(xmlHttpReturnText, "expirations:\[[^\]]*") & "]"), "expirations:", "")
    spot = Replace(Trim(ExtractText(xmlHttpReturnText, "(?:underlying_price:)[\d\.]*")), "underlying_price:", "")
    
    Set lib = New JsonLib_Options
    Set Json = lib.parse(Expirations)
    If TypeName(Json) = "Collection" Then
        ReDim Dates(1 To Json.Count)
        i = 1
        For Each jObj In Json
            temp = jObj.Items
            Dates(i) = DateSerial(temp(0), temp(1), temp(2))
            
            url = "http://www.google.com/finance/option_chain?q=" & Symbol & _
                  "&expd=" & Day(Dates(i)) & "&expm=" & Month(Dates(i)) & "&expy=" & Year(Dates(i)) & "&output=json"
            With xmlHttpRequest
                .Open "GET", url, False
                .send
                xmlHttpReturnText = .responseText
            End With
            Calls = Replace(Trim(ExtractText(xmlHttpReturnText, "calls:\[[^\]]*") & "]"), "calls:", "")
            Puts = Replace(Trim(ExtractText(xmlHttpReturnText, "puts:\[[^\]]*") & "]"), "puts:", "")
            Set JsonOpts = lib.parse(Calls)
            If TypeName(JsonOpts) = "Collection" Then
                ReDim Data(1 To JsonOpts.Count, 1 To 16)
                counter = counter + JsonOpts.Count
                Keys = Array("s", "p", "c", "b", "a", "vol", "oi", "strike", "expiry")
                j = 1
                For Each lObj In JsonOpts
                    l = 1
                    For Each k In Keys
                        Data(j, l) = lObj.Item(k)
                        l = l + 1
                    Next k
                    j = j + 1
                Next lObj
                Set JsonOpts = lib.parse(Puts)
                If TypeName(JsonOpts) = "Collection" Then
                    For Each lObj In JsonOpts
                        For j = 1 To UBound(Data)
                            If Data(j, 9) = lObj.Item("expiry") And Data(j, 8) = lObj.Item("strike") Then
                                l = 9
                                For Each k In Keys
                                    If k <> "expiry" And k <> "strike" Then
                                        Data(j, l) = lObj.Item(k)
                                        l = l + 1
                                    End If
                                Next k
                                If IsDate(Data(j, 1)) Then Data(j, 1) = CDate(Data(j, 1))
                                Exit For
                            End If
                        Next j
                    Next lObj
                End If
                If OptionCollection Is Nothing Then Set OptionCollection = New Collection
                OptionCollection.Add Data
            End If
            i = i + 1
        Next jObj
    End If

    If counter Then
        ReDim FinalData(1 To counter + 1, 1 To 16)
        counter = 1
        For Each Data In OptionCollection
            For i = 1 To UBound(Data)
                counter = counter + 1
                For j = 1 To UBound(Data, 2)
                    FinalData(counter, j) = Data(i, j)
                Next j
            Next i
        Next Data
        FinalData(1, 1) = "Symbol": FinalData(1, 2) = "Last": FinalData(1, 3) = "Change": FinalData(1, 4) = "Bid"
        FinalData(1, 5) = "Ask": FinalData(1, 6) = "Volume": FinalData(1, 7) = "Open Int": FinalData(1, 8) = "Strike"
        For j = 1 To 7
            FinalData(1, j + 8) = FinalData(1, j)
        Next j
        
        Range("gSymbol").Offset(0, 1) = spot
        Range(Range("gSymbol").Offset(2, 0), Range("gSymbol").Offset(UBound(FinalData) + 1, UBound(FinalData, 2) - 1)) = FinalData
        Range(Range("gSymbol").Offset(UBound(FinalData) + 2, 0), _
        Range("gSymbol").Offset(Rows.Count - Range("gSymbol").Row, UBound(FinalData, 2))).ClearContents
        Range(Range("gSymbol").Offset(1, 0), Range("gSymbol").Offset(1, UBound(FinalData, 2) - 1)).ClearContents
    End If

    Set Json = Nothing
    Set JsonOpts = Nothing
    Set lib = Nothing
    Set xmlHttpRequest = Nothing

End Sub


Private Function ExtractText(Expresion As String, Pattern As String) As String
    Dim Rx As RegExp, Rc As Object
    Set Rx = New RegExp
    With Rx
        .IgnoreCase = True
        .Pattern = Pattern
        Set Rc = .Execute(Trim(Expresion))
    End With
    If Rc.Count > 0 Then ExtractText = Trim(Rc.Item(0))
    Set Rc = Nothing
    Set Rx = Nothing
End Function






