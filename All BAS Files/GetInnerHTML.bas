Attribute VB_Name = "GetInnerHTML"
Sub Main_Test()
Dim SH As Worksheet
Set SH = Sheets("Sheet1")
SH.Activate
Dim URLs As Range: Set URLs = Range("B2:B92")
Dim url As Range
For Each url In URLs
    GetInnerHTML (url.Value)
Next url
End Sub
Sub GetInnerHTML(url As String)
With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .Calculation = xlAutomatic
End With

     'we define the essential variables
    Dim ie As Object
    Dim pwd, SearchStr
    Dim button
    Dim xobj As HTMLDivElement
    Dim i As Integer
    Dim trs, tds ', username, pwd, button
    Dim tr As Integer, td As Integer, r As Integer, C As Integer
    'add the "Microsoft Internet Controls" reference in your VBA Project indirectly
    Set ie = CreateObject("InternetExplorer.Application")
    With ie
        .Visible = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        .navigate (url)
        While ie.readyState <> 4
            DoEvents
        Wend
        Debug.Print url & vbNewLine & .document.DocumentElement.innerhtml
        Debug.Print "--------------------------------------------------------"
    End With
    ie.Quit
    Set ie = Nothing
 With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
    .Calculation = xlAutomatic
End With
End Sub

