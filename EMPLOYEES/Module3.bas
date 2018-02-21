Attribute VB_Name = "Module3"
Option Explicit

Sub Load_ReadXMLDoc()
    Dim xmlDoc As MSXML2.DOMDocument60

    ' Create an instance of the DOMDocument
    Set xmlDoc = New MSXML2.DOMDocument60

    ' Disable asynchronous loading
    xmlDoc.async = False

    ' Load XML information from a file
    If xmlDoc.Load("C:\Excel2013_XML\Courses1.xml") Then
        ' Use the DOMDocument object's XML property to
        ' retrieve the raw data
        Debug.Print xmlDoc.XML
        ' Use the DOMDocument object's Text property to
        ' retrieve the actual text stored in nodes
        Sheets.Add
        ActiveSheet.Range("A1").Value = xmlDoc.Text
    End If
End Sub

