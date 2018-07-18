Attribute VB_Name = "Module1"
Sub CreateTOC()
    Dim i As Integer
    Sheets.Add Before:=Sheets(1)
    For i = 2 To Worksheets.Count
      ActiveSheet.Hyperlinks.Add _
         Anchor:=Cells(i, 1), _
         Address:="", _
         SubAddress:="'" & Worksheets(i).Name & "'!A1", _
         TextToDisplay:=Worksheets(i).Name
     Next i
End Sub

