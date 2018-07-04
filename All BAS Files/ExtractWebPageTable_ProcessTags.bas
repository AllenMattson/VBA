4
5
6
7
8
9
10
11
12
13
14
15
16
17
18
19
20
21
22
23
24
25
26
27
28
29
30
31
32
33
34
35
36
37
38
39
40
41
42
43
44
	
Sub HTML_Table_To_Excel()
    Dim htm As Object
    Dim Tr As Object
    Dim Td As Object
    Dim Tab1 As Object
 
    'Replace the URL of the webpage that you want to download
    Web_URL = VBA.Trim(Sheets(1).Cells(1, 1))
 
    'Create HTMLFile Object
    Set HTML_Content = CreateObject("htmlfile")
 
    'Get the WebPage Content to HTMLFile Object
    With CreateObject("msxml2.xmlhttp")
        .Open "GET", Web_URL, False
        .send
        HTML_Content.Body.Innerhtml = .responseText
    End With
 
    Column_Num_To_Start = 1
    iRow = 2
    iCol = Column_Num_To_Start
    iTable = 0
 
    'Loop Through Each Table and Download it to Excel in Proper Format
    For Each Tab1 In HTML_Content.getElementsByTagName("table")
        With HTML_Content.getElementsByTagName("table")(iTable)
            For Each Tr In .Rows
                For Each Td In Tr.Cells
                    Sheets(1).Cells(iRow, iCol).Select
                    Sheets(1).Cells(iRow, iCol) = Td.innerText
                    iCol = iCol + 1
                Next Td
                iCol = Column_Num_To_Start
                iRow = iRow + 1
            Next Tr
        End With
        iTable = iTable + 1
        iCol = Column_Num_To_Start
        iRow = iRow + 1
    Next Tab1
 
    MsgBox "Process Completed"
End Sub