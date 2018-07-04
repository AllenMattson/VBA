Attribute VB_Name = "VBAWebScraping"
'https://raw.githubusercontent.com/HotBreakfast/Web-Scraping-With-VBA/master/rdhVBAWebScraping.bas
Sub Main()
    ' Use the web browser control to retrieve data from web pages
    Dim strResult As String
    Dim strTitle As String
    
    ' Retrieve one web page
    strResult = BasicRetrieve("http://www.google.ca")
    
    ' Display the page title
    strTitle = GetPageTitle(strResult)
    Debug.Print strTitle
    
End Sub


Function GetPageTitle(html As String) As String
' Source: http://analystcave.com/web-scraping-tutorial/
    GetPageTitle = Mid(html, InStr(html, "<title>") _
                 + Len("<title>"), InStr(html, "</title>") _
                 - InStr(html, "<title>") - Len("</title>") + 1)
End Function

Function BasicRetrieve(strURL As String) As String
    ' Source: http://analystcave.com/web-scraping-tutorial/
    
    Dim XMLHTTP As Object
    
    Set XMLHTTP = CreateObject("MSXML2.serverXMLHTTP")
    XMLHTTP.Open "GET", strURL, False
    XMLHTTP.setRequestHeader "Content-Type", "text/xml"
    XMLHTTP.send
    
    ' Return the HTML code the entire page requested
    BasicRetrieve = XMLHTTP.ResponseText
End Function


