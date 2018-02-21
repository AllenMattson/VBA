Attribute VB_Name = "Module8"
Option Explicit

Sub Select_SingleNode_2()
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim xmlSingleN As MSXML2.IXMLDOMNode

    ' Create an instance of the DOMDocument
    Set xmlDoc = New MSXML2.DOMDocument60
    xmlDoc.async = False

    ' Load XML information from a file
    xmlDoc.Load ("C:\Excel2013_XML\Courses1.xml")

    ' Retrieve the reference to a particular node
    Set xmlSingleN = xmlDoc.SelectSingleNode("//Course//@ID")
    If xmlSingleN Is Nothing Then
        Debug.Print "No nodes selected."
    Else
        Debug.Print xmlSingleN.Text
        xmlSingleN.Text = "VBA1EX2013"
        Debug.Print xmlSingleN.Text
        xmlDoc.Save "C:\Excel2013_XML\Courses1.xml"
    End If
End Sub

