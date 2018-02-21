Attribute VB_Name = "Module6"
Option Explicit

Sub SelectNodes_SpecifyCriterion()
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim xmlNodeList As MSXML2.IXMLDOMNodeList
    Dim myNode As Variant

    ' Create an instance of the DOMDocument
    Set xmlDoc = New MSXML2.DOMDocument60
    xmlDoc.async = False

    ' Load XML information from a file
    xmlDoc.Load ("C:\Excel2013_XML\Courses1.xml")

    ' Retrieve all the nodes that match the specified criterion
    Set xmlNodeList = xmlDoc.SelectNodes("//Title")
    If Not (xmlNodeList Is Nothing) Then
        For Each myNode In xmlNodeList
            Debug.Print myNode.Text
        Next myNode
    End If
End Sub

