Attribute VB_Name = "Module7"
Option Explicit

Sub Select_SingleNode()
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim xmlSingleN As MSXML2.IXMLDOMNode

    ' Create an instance of the DOMDocument
    Set xmlDoc = New MSXML2.DOMDocument60
    xmlDoc.async = False

    ' Load XML information from a file
    xmlDoc.Load ("C:\Excel2013_XML\Courses1.xml")

    ' Retrieve the reference to a particular node
    Set xmlSingleN = xmlDoc.SelectSingleNode("//Title")
    Debug.Print xmlSingleN.Text
End Sub

