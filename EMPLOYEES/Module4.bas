Attribute VB_Name = "Module4"
Option Explicit

Sub LearnAboutNodes()
    Dim xmlDoc As MSXML2.DOMDocument60
    Dim xmlNode As MSXML2.IXMLDOMNode

    ' Create an instance of the DOMDocument
    Set xmlDoc = New MSXML2.DOMDocument60

    xmlDoc.async = False

    ' Load XML information from a file
    xmlDoc.Load ("C:\Excel2013_XML\Courses1.xml")

    ' find out the number of child nodes in the document
    If xmlDoc.HasChildNodes Then
        Debug.Print "Number of Child Nodes: " & _
        xmlDoc.ChildNodes.Length

        ' iterate through the child nodes to gather information
        For Each xmlNode In xmlDoc.ChildNodes
            Debug.Print "Node Name: " & xmlNode.nodeName
            Debug.Print vbTab & "Type: " & _
            xmlNode.nodeTypeString & _
            "(" & xmlNode.NodeType & ")"
            Debug.Print vbTab & "Text: " & xmlNode.Text
        Next xmlNode
    End If
End Sub


