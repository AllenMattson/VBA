Attribute VB_Name = "ScrapeMedlineUpdated"
Private MyCount As Long
Sub GetMedlineTableTextofHTML()
Sheets("SKUs").Activate
Sheets("SKUs").Cells.Clear: Sheets("SKUs").Cells(1, 1).Value = "Sku_num": _
Sheets("SKUs").Cells(1, 2).Value = "Description": Sheets("SKUs").Cells(1, 3).Value = "Estimated_Availability": _
Sheets("SKUs").Cells(1, 4).Value = "Packaging": Sheets("SKUs").Cells(1, 5).Value = "QTY": _
Sheets("SKUs").Cells(1, 6).Value = "Price": Sheets("SKUs").Cells(1, 7).Value = "Contract_Price"
Sheets("SKUs").Cells(1, 6).Value = "Price": Sheets("SKUs").Cells(1, 8).Value = "URL_ID"

MyCount = 1

    Dim ObjIE As Object
    Dim pwd, username
    Dim button

    'add the "Microsoft Internet Controls" reference in your VBA Project indirectly
    Set ObjIE = CreateObject("InternetExplorer.Application")
    With ObjIE
        .Visible = True
        .navigate ("https://www.medline.com/account/login.jsp?fromst=true#")
        While ObjIE.readyState <> 4
            DoEvents
        Wend
        Set username = .document.getElementById("ext-gen1004") 'id of the username control (HTML Control)'ext-gen1004
        Set pwd = .document.getElementById("ext-gen1005") 'id of the password control (HTML Control) 'ext-gen1005
        Set button = .document.getElementById("submitbutton") 'id of the button control (HTML Control) 'submitbutton
        username.Value = "danya"
        pwd.Value = "Geritom10"
        button.Click
        While ObjIE.readyState <> 4
            DoEvents
        Wend
    End With
    Application.Wait Now + TimeValue("0:00:05")
    ''''''''''''''''''''''''''''


    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With


Dim SH As Worksheet: Set SH = Sheets("Sheet1"): SH.Activate
Dim LR As Integer: LR = SH.Cells(Rows.Count, 1).End(xlUp).Row

Dim MyRNG As Range: Set MyRNG = SH.Range("A2:A" & LR)

'Module1.IESignIn

    
    Dim xobj As HTMLDivElement
    Dim i As Integer
    Dim trs, tds ', username, pwd, button
    Dim tr As Integer, td As Integer, r As Integer, C As Integer
    
    


Dim Cell As Range
Dim URL As String


Dim SKUsh As Worksheet: Set SKUsh = Sheets("SKUs")
Dim LastRo As Long: LastRo = SKUsh.Cells(Rows.Count, 1).End(xlUp).Row


With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .Calculation = xlAutomatic
End With

With SH
'If objie Is Nothing Then Set objie = CreateObject("InternetExplorer.Application"):

DoEvents
For Each Cell In MyRNG
'create unique url id for debug testing and key id to SKUs sheet
    With SH
        If Cell.Value <> "" Then
            Cell.Offset(0, 4).Value = MyCount
            MyCount = MyCount + 1
        End If
    End With
    URL = Cell.Value
    SKUsh.Activate: SKUsh.Cells(LastRo, 1).Offset(1, 0).Activate: DoEvents
    With ObjIE
        .Visible = True
        .navigate URL
    End With

    'establish connection to internet, catch if IE crashes
    On Error GoTo ProductUnavailable: While (ObjIE.Busy Or ObjIE.readyState <> 4): Application.Wait Now + TimeValue("0:00:05"): Wend
    
    'Use top div id name that holds table, move down elements by classes
    Set xobj = ObjIE.document.getElementById("productFamilyWpOrderingInfo") ''''"myDiv"'''''
    Set xobj = xobj.getElementsByClassName("medSkuProductListExpanded productFamilyOrderInfo").Item(0) ''''"myTable"''''
    Set xobj = xobj.getElementsByClassName("medGridViewSkuList persist-area")(0) ''''"data"(0)'''''
    If xobj Is Nothing = True Then GoTo ProductUnavailable
    
    'Grab the table with the info needed, identify rows and columns to extract data into spreadsheet
    Set xobj = xobj.getElementsByClassName("actualDataTable")(0)
    Set trs = xobj.getElementsByClassName("skuRow    ")
    
    'Loop through table rows and for each column insert data into excel
    'Currently the data is inserted bottom up
    With SKUsh
        For r = 0 To trs.Length - 1
            Set tds = trs(r).getElementsByTagName("td")
                For C = 0 To tds.Length - 1
                    SKUsh.Cells(100000, C + 1).End(xlUp).Offset(1, 0).Value = tds(C).innerText
                Next C
                SKUsh.Cells(100000, 1).End(xlUp).Offset(1, 6).Value = MyCount
        Next r
    End With
    
ProductUnavailable:
        If Err.Number = 462 Then
            Debug.Print "ERROR IE JUST BIT THE DUST" & vbNewLine & "LAST URL: " & URL
        Else
            If Err.Number > 0 Or Err.Number < 0 Then
                Debug.Print "Error: " & Err.Number & vbNewLine & URL & " This product is currently unavailable"
            End If
        End If
Next Cell

'wipe the url table variable clean
Set xobj = Nothing

'objIe.Quit
'remove object created for variable
Set ObjIE = Nothing
End With
'LEAVE APP LEVEL EVENTS ON? doesnt hurt i guess...
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



