Attribute VB_Name = "ProductSearch"
Sub SearchProduct()
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .Calculation = xlAutomatic
End With
Dim Cell As Range
Dim MyProductList As Range: Set MyProductList = Sheets("SKUs").Range("A2:A764")


     'we define the essential variables
    Dim ie As Object
    Dim pwd, SearchStr
    Dim button
    Dim xobj As HTMLDivElement
    Dim i As Integer
    Dim trs, tds ', username, pwd, button
    Dim tr As Integer, td As Integer, r As Integer, C As Integer
    'add the "Microsoft Internet Controls" reference in your VBA Project indirectly
    Set ie = CreateObject("InternetExplorer.Application")
    With ie
        .Visible = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        .navigate ("https://www.medline.com/account/login.jsp?fromst=true#")
        While ie.readyState <> 4
            DoEvents
        Wend
        Set username = .document.getElementById("ext-gen1004") 'id of the username control (HTML Control)'ext-gen1004
        Set pwd = .document.getElementById("ext-gen1005") 'id of the password control (HTML Control) 'ext-gen1005
        Set button = .document.getElementById("submitbutton") 'id of the button control (HTML Control) 'submitbutton
        username.Value = "danya"
        pwd.Value = "Geritom10"
        button.Click
        While ie.readyState <> 4
            DoEvents
        Wend
        Application.Wait Now + TimeValue("0:00:05")
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For Each Cell In MyProductList
            Set SearchStr = .document.getElementById("searchQuestion") 'id of the username control (HTML Control)
            Set button = .document.getElementById("searchSubmit") 'id of the button control (HTML Control) 'searchSubmit
            'Search For Item
            SearchStr.Value = Cell.Value
            button.Click
            While ie.readyState <> 4
                DoEvents
            Wend
            Application.Wait Now + TimeValue("0:00:05"): On Error GoTo ErrHandle
            Set xobj = ie.document.getElementById("galleryView") ''''"myDiv"'''''
            Set xobj = xobj.getElementsByClassName("medGridViewSkuListWrapper").Item(0)
            Set xobj = xobj.getElementsByClassName("medGridViewSkuList persist-area")(0)
            Set trs = xobj.getElementsByClassName("skuRow    ")
            'Loop Table
            For r = 0 To trs.Length - 1
                Set tds = trs(r).getElementsByTagName("td")
                    For C = 0 To tds.Length - 1
                    'remove last ordered row by making length>2
                        If tds.Length > 2 Then
                            Sheets("Sheet1").Cells(100000, C + 1).End(xlUp).Offset(1, 0).Value = tds(C).innerText
                        End If
                    Next C
            Next r
ErrHandle:
            Application.Wait Now + TimeValue("0:00:03")
            Next Cell
    End With
    ie.Quit
    Set ie = Nothing
    
End Sub

