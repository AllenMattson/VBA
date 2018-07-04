Attribute VB_Name = "ReadXMLintoExcel"
'In Tools > References, add reference to "Microsoft XML, vX.X" before running.
Sub subReadXMLStream()
    
    Dim xmlDoc As MSXML2.DOMDocument
    Dim xEmpDetails As MSXML2.IXMLDOMNode
    Dim xParent As MSXML2.IXMLDOMNode
    Dim xChild As MSXML2.IXMLDOMNode
    Dim Col, Row As Integer

    Set xmlDoc = New MSXML2.DOMDocument
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    ' use XML string to create a DOM, on error show error message
    If Not xmlDoc.Load("http://itpscan.info/blog/excel/xml/schedule.xml") Then
        Err.Raise xmlDoc.parseError.ErrorCode, , xmlDoc.parseError.reason
    End If
        
    Set xEmpDetails = xmlDoc.DocumentElement
    Set xParent = xEmpDetails.FirstChild
    
    Row = 1
    Col = 1
    
    Dim xmlNodeList As IXMLDOMNodeList
    
    Set xmlNodeList = xmlDoc.SelectNodes("//record")
    
    For Each xParent In xmlNodeList
        For Each xChild In xParent.ChildNodes
            Worksheets("Sheet1").Cells(Row, Col).Value = xChild.Text
            Debug.Print Row & " - "; Col & " -  " & xChild.Text
            Col = Col + 1
        Next xChild
        Row = Row + 1
        Col = 1
    Next xParent
End Sub
