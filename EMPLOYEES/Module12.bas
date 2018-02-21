Attribute VB_Name = "Module12"
Sub Load_ReadXMLDoc_FilterXML()
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim retval As String
    
    ' Create an instance of the DOMDocument60
    Set xmlDoc = New MSXML2.DOMDocument60

    ' Disable asynchronous loading
    xmlDoc.async = False

    ' Load XML information from a file
    If xmlDoc.Load("C:\Excel2013_XML\Courses1.xml") Then
        ' Use the DOMDocument60 object's XML property to
        ' retrieve the raw data to the worksheet
        Sheets.Add
        ActiveSheet.Range("A1").Value = xmlDoc.XML
        Columns("A:A").ColumnWidth = 65
        ' Use the Excel 2013 built-in function FilterXML to
        ' retrieve data stored in a specific node
        ActiveSheet.Range("A4").Value = _
          WorksheetFunction.FilterXML( _
          Range("A1").Value, "//Course[@ID='VBA2EX']//Title")
    End If
End Sub


