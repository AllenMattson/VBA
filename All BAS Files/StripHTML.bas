Attribute VB_Name = "StripHTML"
Sub RetrieveHTML()
     
     
    Dim rngSelect As Range
    Dim sURL As String
     
     
    Set rngSelect = Range("A1", Range("A1").End(xlDown))
    'Debug.Print rngSelect.Address
     
     
    Set IE = CreateObject("InternetExplorer.Application")
     
     
    For Each C In rngSelect
        sURL = C.Value
        If sURL = "" Then GoTo ImDoneNow
        With IE
            .Visible = False 'True 'You can set this to False and it won't show you what it is doing
            .Navigate sURL
            Do Until .READYSTATE = 4
                DoEvents
            Loop
            Do While .Busy: DoEvents: Loop
                'Range(C.Address).Offset(0, 1).Value = IE.Document.DocumentElement.innerHTML 'OuterHTML
                If Len(Trim(IE.Document.DocumentElement.innerHTML)) > 0 Then
                    WriteUpInnerHTML (IE.Document.DocumentElement.innerHTML)
                Else
                    If Len(Trim(IE.Document.DocumentElement.innerHTML)) = 0 Then GoTo ImDoneNow
                End If
            End With
        Next C
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ImDoneNow:
Dim MyName As String
MyName = Cells(1, 1).Value
'Pull out site name for naming convention
If InStr(MyName, "www.") Then
    nameStart = InStr(MyName, "www.") + 3 'Add 3 because the integer pulled is from the first w in www.
    MyName = Right(MyName, Len(MyName) - nameStart)
    MyName = Left(MyName, InStr(MyName, ".com") - 1) 'subtract 1 to get rid of the . in the ".com"
    'Debug.Print MyName
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
MsgBox "A file named " & MyName & ".txt has been placed into the following directory: " & vbNewLine _
    & vbNewLine & Application.ThisWorkbook.Path, vbOKOnly, "Saved HTML into text file..."
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
Sub WriteUpInnerHTML(MyData As String)
Dim AllTheHTML As String
Dim FilePath As String, CellData As String, MyName As String
Dim i As Integer, j As Integer, nameStart As Integer
Dim LC As Long, LR As Long
'LC = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
'LR = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
LR = Cells(Rows.Count, 2).End(xlUp).Row
'We have all html in column "B" therefore LC is 2 for now
LC = 2

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
MyName = Cells(1, 1).Value
'Pull out site name for naming convention
If InStr(MyName, "www.") Then
    nameStart = InStr(MyName, "www.") + 3 'Add 3 because the integer pulled is from the first w in www.
    MyName = Right(MyName, Len(MyName) - nameStart)
    MyName = Left(MyName, InStr(MyName, ".com") - 1) 'subtract 1 to get rid of the . in the ".com"
    'Debug.Print MyName
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

CellData = ""
FilePath = Application.ThisWorkbook.Path & "\" & MyName & ".txt"

Open FilePath For Output As #2
CellData = MyData
Write #2, CellData
CellData = ""
Close #2


End Sub
Function StripHTML(cell As Range) As String
    Dim RegEx As Object
    Set RegEx = CreateObject("vbscript.regexp")
     
     
    Dim sInput As String
    Dim sOut As String
    sInput = cell.Text
     
     
    With RegEx
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "<[^>]+>" 'Regular Expression for HTML Tags.
    End With
     
     
    sOut = RegEx.Replace(sInput, "")
    StripHTML = sOut
    Set RegEx = Nothing
End Function
