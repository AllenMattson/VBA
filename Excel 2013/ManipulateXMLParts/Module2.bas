Attribute VB_Name = "Module2"
Sub ListUniqueValues()
    Dim xmlDoc As DOMDocument60
    Dim myNodeList As IXMLDOMNodeList
    Dim i As Integer
    Dim iLen As Integer

    Set xmlDoc = New DOMDocument60
    xmlDoc.async = False
    xmlDoc.Load ("C:\Excel2013_ByExample\ZipPackage\xl\sharedStrings.xml")
    Set myNodeList = xmlDoc.SelectNodes("//t")
    iLen = myNodeList.Length

    Worksheets(1).Activate
    For i = 0 To iLen - 1
        Range("A" & i + 1).Formula = myNodeList(i).Text
    Next
    Columns("A").AutoFit

    Set myNodeList = Nothing
    Set xmlDoc = Nothing
End Sub

