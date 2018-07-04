Attribute VB_Name = "MedlineScrape_WORKS"
Private MyCount As Long
Sub GetMedlineTableTextofHTML()
MyErrorHandle:
'Sheets("SKUs").Activate
'Sheets("SKUs").Cells.Clear: Sheets("SKUs").Cells(1, 1).Value = "Sku_num": _
'Sheets("SKUs").Cells(1, 2).Value = "Description": Sheets("SKUs").Cells(1, 3).Value = "Estimated_Availability": _
'Sheets("SKUs").Cells(1, 4).Value = "Packaging": Sheets("SKUs").Cells(1, 5).Value = "QTY": _
'Sheets("SKUs").Cells(1, 6).Value = "Price": Sheets("SKUs").Cells(1, 7).Value = "Contract_Price"
'Sheets("SKUs").Cells(1, 6).Value = "Price": Sheets("SKUs").Cells(1, 8).Value = "URL_ID"
Dim SH As Worksheet: Set SH = Sheets("XML_SiteMap"): SH.Activate
Dim LR As Long: LR = SH.Cells(Rows.Count, 1).End(xlUp).Row
'establish range to loop through

Dim MyCountNum As Long: MyCountNum = SH.Cells(Rows.Count, "E").End(xlUp).Row
MyCount = MyCountNum
Dim MyRNG As Range: Set MyRNG = SH.Range("A" & MyCountNum + 1 & ":A" & LR)
'establish counter


    Dim ObjIE As Object
    Dim pwd, username
    Dim button


    

'Module1.IESignIn

    
    Dim Xobj As HTMLDivElement
    Dim i As Integer
    Dim trs, tds ', username, pwd, button
    Dim tr As Integer, td As Integer, r As Integer, C As Integer
    
    


Dim Cell As Range
Dim URL As String


Dim SKUsh As Worksheet: Set SKUsh = Sheets("SKUs")
Dim LastRo As Long: LastRo = SKUsh.Cells(Rows.Count, 1).End(xlUp).Row


With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .Calculation = xlManual
End With

With SH
'add the "Microsoft Internet Controls" reference in your VBA Project indirectly
    Set ObjIE = CreateObject("InternetExplorer.Application")
For Each Cell In MyRNG
'create unique url id for debug testing and key id to SKUs sheet
    With SH
        If Cell.Value <> "" Then
            Cell.Offset(0, 4).Value = MyCount
            MyCount = MyCount + 1
        End If
    End With
Loop_Begin:
    URL = Cell.Value
    SKUsh.Activate: SKUsh.Cells(LastRo, 1).Offset(1, 0).Activate: DoEvents
    With ObjIE
        .Visible = False 'True
        .navigate URL
    End With

    'establish connection to internet, catch if IE crashes
    On Error GoTo MyErrorHandle: While (ObjIE.Busy Or ObjIE.readyState <> 4): Application.Wait Now + TimeValue("0:00:01"): Wend
    If ObjIE Is Nothing Then GoTo MyErrorHandle
        Set Xobj = ObjIE.document.getElementById("productFamilyWpOrderingInfo") ''''"myDiv"'''''
        Set Xobj = Xobj.getElementsByClassName("medSkuProductListExpanded productFamilyOrderInfo").Item(0) ''''"myTable"''''
        Set Xobj = Xobj.getElementsByClassName("medGridViewSkuList persist-area")(0) ''''"data"(0)'''''
        If Not Xobj Is Nothing Then
            Set Xobj = Xobj.getElementsByClassName("actualDataTable")(0)
        Else
            GoTo MyErrorHandle
        End If
        Set trs = Xobj.getElementsByClassName("skuRow    ")

            With SKUsh
                For r = 0 To trs.Length - 1
                    Set tds = trs(r).getElementsByTagName("td")
                        For C = 0 To tds.Length - 1
                            SKUsh.Cells(100000, C + 1).End(xlUp).Offset(1, 0).Value = tds(C).innerText
                        Next C
                        SKUsh.Cells(100000, 1).End(xlUp).Offset(0, 3).Value = URL
                Next r
            End With

ProductUnavailable:
        If Err.Number = 462 Then
            Debug.Print "ERROR IE JUST BIT THE DUST" & vbNewLine & "LAST URL: " & URL
        Else
                If Err.Number > 0 Or Err.Number < 0 Then
                    Debug.Print "Error: " & Err.Number & " Description: " & Err.Description & " "; URL & " This product is currently unavailable"
                    If Not ObjIE Is Nothing Then
                        On Error Resume Next: ObjIE.Quit: Set ObjIE = Nothing: Set Xobj = Nothing: DoEvents: GoTo MyErrorHandle
                    Else
                        GoTo MyErrorHandle
                    End If
                End If
        End If
        Debug.Print MyCount & " URL: " & URL & " has been scraped."
'wipe the url table variable clean
Set Xobj = Nothing

Next Cell


ObjIE.Quit
'remove object created for variable
Set ObjIE = Nothing
End With
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .Calculation = xlAutomatic
End With
End Sub
Sub IESignIn()
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .Calculation = xlAutomatic
End With
     'we define the essential variables
    Dim ie As Object
    Dim pwd, username
    Dim button

    'add the "Microsoft Internet Controls" reference in your VBA Project indirectly
    Set ie = CreateObject("InternetExplorer.Application")
    With ie
        .Visible = True
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
    End With
    Set ie = Nothing
    Application.Wait Now + TimeValue("0:00:05")
End Sub


Sub SearchProduct()
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .Calculation = xlAutomatic
End With
     'we define the essential variables
    Dim ie As Object
    Dim pwd, SearchStr
    Dim button
    Dim Xobj As HTMLDivElement
    Dim i As Integer
    Dim trs, tds ', username, pwd, button
    Dim tr As Integer, td As Integer, r As Integer, C As Integer
    'add the "Microsoft Internet Controls" reference in your VBA Project indirectly
    Set ie = CreateObject("InternetExplorer.Application")
    With ie
        .Visible = True
        ''''''''''''''''''''''
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
        Application.Wait Now + TimeValue("0:00:02")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Set SearchStr = .document.getElementById("searchQuestion") 'id of the username control (HTML Control)'ext-gen1004
        Set button = .document.getElementById("searchSubmit") '("button-blue search right") 'id of the button control (HTML Control) 'submitbutton
        SearchStr.Value = "COI21053"
        button.Click
        While ie.readyState <> 4
            DoEvents
        Wend
     Application.Wait Now + TimeValue("0:00:02")
    Set Xobj = ie.document.getElementById("galleryView") ''''"myDiv"'''''
    Set Xobj = Xobj.getElementsByClassName("medGridViewSkuListWrapper").Item(0)
    Set Xobj = Xobj.getElementsByClassName("medGridViewSkuList persist-area")(0)
        Set trs = Xobj.getElementsByClassName("skuRow    ")
        For r = 0 To trs.Length - 1
            Set tds = trs(r).getElementsByTagName("td")
                For C = 0 To tds.Length - 1
                    Debug.Print tds(C).innerText 'SKUsh.Cells(100000, C + 1).End(xlUp).Offset(1, 0).Value = tds(C).innerText
                Next C
        Next r
    End With
    Set ie = Nothing
    
End Sub


