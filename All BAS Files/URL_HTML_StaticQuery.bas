Attribute VB_Name = "URL_HTML_StaticQuery"
Sub URL_HTML_StaticQuery()
Dim MyStr As String
MyStr = Cells(1, 1).Value
If Left(MyStr, 4) = "http" Then MyStr = Right(MyStr, Len(MyStr) - 4)

IND_URL_Static_Query (MyStr)
End Sub
Private Function IND_URL_Static_Query(cell As String)
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual
On Error Resume Next
   With ActiveSheet.QueryTables.Add(Connection:= _
      "URL;http://" & cell, _
         Destination:=Range("a2"))
      .BackgroundQuery = True
      .TablesOnlyFromHTML = True
      .Refresh BackgroundQuery:=False
      .SaveData = True
   End With
ActiveSheet.Range("a1").CurrentRegion.TextToColumns Destination:=ActiveSheet.Range("a1"), DataType:=xlDelimited, _
TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
Semicolon:=False, Comma:=True, Space:=False, other:=False
ActiveSheet.Columns.AutoFit
End Function

