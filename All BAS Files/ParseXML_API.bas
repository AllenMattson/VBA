Attribute VB_Name = "ParseXML_API"
Sub GetLibor()
'==================================================================================
'API KEY: 9047e47c2e1112d062a909fd86472033
'Name a range [APIurl]
'Enter this into named range: https://research.stlouisfed.org/useraccount/apikey?api_key=9047e47c2e1112d062a909fd86472033
'Video: https://www.youtube.com/results?search_query=vba+twitter+api
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''References Needed to turn on''''
'microsoft internet controls
'microsoft winhttp services 5.1
'microsoft xml 6.0
'==================================================================================
Dim ws As Worksheet: Set ws = Worksheets("API")
Dim strURL As String: strURL = ws.[APIurl]

Dim hReq As New WinHttpRequest
hReq.Open "GET", strURL, False
hReq.send

Dim strRESP As String
strRESP = hReq.responseText

Dim xmlDOC As New MSXML2.DOMDocument60
If Not xmlDOC.LoadXML(strRESP) Then
    MsgBox "Load Error"
End If


'Get nodes
Dim xNodeList As MSXML2.IXMLDOMNodeList: Set xNodeList = xmlDOC.getElementsByTagName("Observations")
'Define Node
Dim xNode As MSXML2.IXMLDOMNode: Set xNode = xNodeList.Item(0)
'Attributes to obtain
Dim obAtt1 As MSXML2.IXMLDOMAttribute
Dim obAtt2 As MSXML2.IXMLDOMAttribute


Dim xChild As MSXML2.IXMLDOMNode

Dim intRow As Integer: intRow = 2
Dim strCol1 As String: strCol1 = "A"
Dim strCol2 As String: strCol1 = "B"


Dim dtVar As Date
Dim dblRate As Double
Dim strVal As String


For Each xChild In xNode.ChildNodes
    Set obAtt1 = xChild.Attributes.getNamedItem("date")
    Set obAtt2 = xChild.Attributes.getNamedItem("value")
    strVal = Trim(obAtt2.Text)
    If strVal = "." Then
        ws.Cells(intRow, 2) = ""
    Else
        ws.Cells(intRow, 2) = Format(strVal / 100, "0.00%") ' strVal
    End If
    ws.Cells(intRow, 1) = CDate(Trim(obAtt1.Text))
    intRow = intRow + 1
Next xChild

Set hReq = Nothing
Set xmlDOC = Nothing

End Sub
