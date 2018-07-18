Attribute VB_Name = "Module5"
Option Explicit

Sub ChangeLeftMargin_RemovePageSetup()
    Dim xmlDoc As DOMDocument60
    Dim myNode As Msxml2.IXMLDOMNode
    Dim strSrchNode As String

    
    Set xmlDoc = New DOMDocument60
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    xmlDoc.Load ("C:\Excel2013_ByExample\ZipPackage\xl\" _
        & "worksheets\Sheet1.XML")
    
    xmlDoc.setProperty "SelectionNamespaces", _
        "xmlns:x14ac='http://schemas.openxmlformats.org/" & _
            "spreadsheetml/2006/main'"

    strSrchNode = "/x14ac:worksheet/x14ac:pageMargins/@left"

    Set myNode = xmlDoc.selectSingleNode(strSrchNode)

    Debug.Print "previous left margin = " & myNode.Text
    
    myNode.Text = "0.50"
    
    Set myNode = xmlDoc.selectSingleNode("//x14ac:pageSetup")
    
    On Error Resume Next
    myNode.ParentNode.RemoveChild myNode
    xmlDoc.Save ("C:\Excel2013_ByExample\ZipPackage\xl\" _
        & "worksheets\Sheet1.XML ")

    Set myNode = Nothing
    Set xmlDoc = Nothing
End Sub


