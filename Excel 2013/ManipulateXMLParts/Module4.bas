Attribute VB_Name = "Module4"
Sub RetrieveAllTextValues()
    Dim xmlDoc As DOMDocument60
    Dim myNodeList1 As IXMLDOMNodeList
    Dim myNodeList2 As IXMLDOMNodeList
    Dim myNodeList3 As IXMLDOMNodeList
    Dim strArray() As String
    Dim i As Integer
    Dim iLen As Integer

    Set xmlDoc = New DOMDocument60
    xmlDoc.async = False

    xmlDoc.Load ("C:\Excel2013_ByExample\ZipPackage\xl\sharedStrings.xml")
    Set myNodeList1 = xmlDoc.SelectNodes("//t")

    iLen = myNodeList1.Length
    ReDim strArray(iLen)

    For i = 0 To iLen - 1
        strArray(i) = myNodeList1(i).Text
    Next

    xmlDoc.async = False
    xmlDoc.Load ("C:\Excel2013_ByExample\ZipPackage\xl\worksheets\sheet1.xml")
    Set myNodeList2 = xmlDoc.SelectNodes("//sheetData/row/c[@t='s']/@r")
    Set myNodeList3 = xmlDoc.SelectNodes("//sheetData/row/c[@t='s']/v")

    Sheets.Add
    'Worksheets(2).Activate
    i = 0

    For i = 0 To myNodeList2.Length - 1
        With Range(myNodeList2(i).Text)
            .Value = strArray(myNodeList3(i).Text)
        End With
    Next

    Range("A1").CurrentRegion.Select
    Selection.EntireColumn.AutoFit

    Set myNodeList1 = Nothing
    Set myNodeList2 = Nothing
    Set myNodeList3 = Nothing
    Set xmlDoc = Nothing
End Sub

