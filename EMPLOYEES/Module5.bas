Attribute VB_Name = "Module5"
Option Explicit

Sub IterateThruElements()
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim xmlNodeList As MSXML2.IXMLDOMNodeList
    Dim xmlNode As MSXML2.IXMLDOMNode
    Dim myNode As MSXML2.IXMLDOMNode

    ' Create an instance of the DOMDocument
    Set xmlDoc = New MSXML2.DOMDocument60
    xmlDoc.async = False

    ' Load XML information from a file
    xmlDoc.Load ("C:\Excel2013_XML\Courses1.xml")

    ' Find out the number of child nodes in the document
    Set xmlNodeList = xmlDoc.getElementsByTagName("*")

    ' Open a new workbook and paste the data
    Workbooks.Add
    Range("A1:B1").Formula = Array("Element Name", "Text")
    For Each xmlNode In xmlNodeList
        For Each myNode In xmlNode.ChildNodes
            If myNode.NodeType = NODE_TEXT Then
                ActiveCell.Offset(0, 0).Formula = xmlNode.nodeName
                ActiveCell.Offset(0, 1).Formula = xmlNode.Text
            End If
        Next myNode
        ActiveCell.Offset(1, 0).Select
    Next xmlNode
    Columns("A:B").AutoFit
End Sub

