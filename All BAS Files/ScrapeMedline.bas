Attribute VB_Name = "ScrapeMedline"
Option Explicit
Sub MainLoop()
Sheets("SKUs").Activate
Sheets("SKUs").Cells.Clear: Sheets("SKUs").Cells(1, 1).Value = "Sku_num": Sheets("SKUs").Cells(1, 2).Value = "Description": Sheets("SKUs").Cells(1, 3).Value = "Quantity"
DoEvents
With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
    .Calculation = xlCalculationManual
End With

'loop the urls retrieved from xml file in hidden sitemap
Dim WS As Worksheet: Set WS = Sheets("Sheet1")
Dim i As Integer
Dim ProductURL As String
With WS
    For i = 71 To 31339
        ProductURL = WS.Cells(i, 1)
        GetInnerTextofHTML (ProductURL)
    Next i
End With
        
End Sub
'set reference to Microsoft HTML Object Library
Sub GetInnerTextofHTML(URL As String)
Sheets("SKUs").Range("A1").End(xlDown).Activate
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .Wait Now + TimeValue("0:00:04")
End With
Sheets("SKUs").Range("A1").End(xlDown).Select
Dim MyProductSkuWS As Worksheet: Set MyProductSkuWS = Sheets("SKUs")
Dim objIe As Object, xobj As HTMLDivElement
Dim i As Integer
Dim trs, tds
Dim tr As Integer, td As Integer, r As Integer, c As Integer

Set objIe = CreateObject("InternetExplorer.Application"): DoEvents
With objIe
    .Visible = False
    .navigate URL
End With

'establish connection to internet, catch if IE crashes
On Error GoTo ProductUnavailable: While (objIe.Busy Or objIe.READYSTATE <> 4): Application.Wait Now + TimeValue("0:00:01"): Wend

'Use top div id name that holds table, move down elements by classes
Set xobj = objIe.document.getElementById("productFamilyWpOrderingInfo") ''''"myDiv"'''''
Set xobj = xobj.getElementsByClassName("medSkuProductListExpanded productFamilyOrderInfo").Item(0) ''''"myTable"''''
Set xobj = xobj.getElementsByClassName("medGridViewSkuList persist-area")(0) ''''"data"(0)'''''

'If product is out of stock, there is not table to query
'debug a message for user and move to the next URL
If xobj Is Nothing = True Then GoTo ProductUnavailable

'Grab the table with the info needed, identify rows and columns to extract data into spreadsheet
Set xobj = xobj.getElementsByClassName("actualDataTable")(0)
Set trs = xobj.getElementsByClassName("skuRow    ")

'Loop through table rows and for each column insert data into excel
'Currently the data is inserted bottom up
With MyProductSkuWS
    For r = 0 To trs.Length - 1
        Set tds = trs(r).getElementsByTagName("td")
            For c = 0 To tds.Length - 1
                MyProductSkuWS.Cells(100000, c + 1).End(xlUp).Offset(1, 0).Value = tds(c).innerText
            Next c
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
    
'wipe the url table variable clean
Set xobj = Nothing

'If IE did not crash then quit IE (otherwise application error)
If objIe <> "" Then objIe.Quit

'remove object created for variable
Set objIe = Nothing

'LEAVE APP LEVEL EVENTS ON? doesnt hurt i guess...
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .Calculation = xlAutomatic
End With
End Sub
