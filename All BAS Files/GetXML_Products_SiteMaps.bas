Attribute VB_Name = "GetXML_Products_SiteMaps"
Sub GetXML_Products_SiteMaps()
Dim WB As Workbook: Set WB = ThisWorkbook
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
                    .CommandType = 0
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
            'Wait for import...(possibly turn background refresh to true)
            Application.Wait (Now + TimeValue("00:00:10"))
        Next i
End With
End Sub
