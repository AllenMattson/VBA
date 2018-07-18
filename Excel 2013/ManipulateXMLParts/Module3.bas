Attribute VB_Name = "Module3"
Sub Text_Replace()
    Dim xmlDoc As DOMDocument60
    Dim myNode As IXMLDOMNode
    Dim srchStr As String
    Dim newStr As String
    Dim strFileToEdit As String

    strFileToEdit = "C:\Excel2013_ByExample\ZipPackage\xl\sharedStrings.xml"

    Call UnzipExcelFile
    If blnIsFileSelected = False Then Exit Sub

    Set xmlDoc = New DOMDocument60
    xmlDoc.async = False
    xmlDoc.Load (strFileToEdit)

    srchStr = InputBox("Please enter the string to find:", _
    "Search for String")

    If srchStr <> "" Then
        ' find the text that needs to be replaced
        Set myNode = xmlDoc.selectSingleNode("//t[text()='" + _
        srchStr + "']")
        If myNode Is Nothing Then Exit Sub
    Else
        Exit Sub
    End If

    ' replace text
    newStr = InputBox("Please enter the replacement string for " _
    & srchStr, "Replace with String")
    If newStr <> "" Then
        myNode.Text = newStr
        xmlDoc.Save strFileToEdit
    Else
        Exit Sub
    End If

    ' zip the files in the package
    Call ZipToExcel

    Set xmlDoc = Nothing
    Set myNode = Nothing

End Sub


