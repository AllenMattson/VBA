Attribute VB_Name = "GetXML_Products_SiteMaps_2"
Sub GetXML_Products_SiteMaps()
With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
End With

Dim WB As Workbook: Set WB = Workbooks.Add 'ThisWorkbook
Dim NewSH As Worksheet
Dim MyXML_URLs(0 To 5) As String
MyXML_URLs(1) = "http://www.medline.com/product.xml"
MyXML_URLs(2) = "http://www.medline.com/product2.xml"
MyXML_URLs(3) = "http://www.medline.com/product3.xml"
MyXML_URLs(4) = "http://www.medline.com/product4.xml"
MyXML_URLs(5) = "http://www.medline.com/product5.xml"

Dim NumCount As Integer
Dim My_XML_URL As String
With WB
    For NumCount = 1 To 5
        My_XML_URL = MyXML_URLs(NumCount)
        Set NewSH = .Worksheets.Add
            With NewSH.QueryTables.Add(Connection:="FINDER;" & My_XML_URL, Destination:=Range("$A$1"))
                'ALL OPTIONS FOR QUERY TABLES SOME MIGHT NOT BE NEEDED BUT ARE LISTED
'                    .CommandType = 0
                    .Name = My_XML_URL
                    .FieldNames = True
                    .RowNumbers = False
                    .FillAdjacentFormulas = False
                    .PreserveFormatting = True
                    .RefreshOnFileOpen = False
                    .BackgroundQuery = False
                    .RefreshStyle = xlInsertDeleteCells
                    .SavePassword = False
                    .SaveData = True
                    .AdjustColumnWidth = True
                    .RefreshPeriod = 0
                    .WebSelectionType = xlAllTables
                    .WebFormatting = xlWebFormattingNone
                    .WebPreFormattedTextToColumns = True
                    .WebConsecutiveDelimitersAsOne = True
                    .WebSingleBlockTextImport = False
                    .WebDisableDateRecognition = False
                    .WebDisableRedirections = False
                    .Refresh BackgroundQuery:=False
            End With
        Next NumCount
End With
DoEvents

Application.Calculate: Application.ScreenUpdating = True
Application.Wait (Now + TimeValue("00:00:05"))
Application.Calculate: Application.ScreenUpdating = False

Application.DisplayAlerts = False
On Error Resume Next
Sheets("XML_SiteMap").Delete
'COMBINE RESULTS INTO SINGLE TABLE
With WB
Dim Master As Worksheet: Set Master = WB.Sheets.Add: ActiveSheet.Name = "XML_SiteMap"
    Dim SH_Done As Worksheet
    Dim SH As Worksheet
    For Each SH In WB.Worksheets
        Application.CutCopyMode = False
        If Not SH.Name = Master.Name Then
            If Trim(SH.Range("B2").Value) = "/url/loc" Then
                Dim LastRo As Integer
                LastRo = SH.Cells(Rows.Count, 2).End(xlUp).Row
                SH.Range("B3:B" & LastRo).Copy Master.Range("A65536").End(xlUp)(2)
            End If
        End If
    Next SH
End With
Application.CutCopyMode = False
Dim XMLwb As Workbook: Set XMLwb = Sheets("XML_SiteMap").Move ' Before:=Workbooks("MEDLINE.xlsm").Sheets(1)
DoEvents
ActiveSheet.Cells.Copy
ActiveSheet.PasteSpecial xlpastvalues

WB.Close False
Application.DisplayAlerts = True
With Application
    .ScreenUpdating = True
    .Calculation = xlAutomatic
End With
End Sub
