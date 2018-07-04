Attribute VB_Name = "AssetCorrelation"
Option Explicit
Sub StartCorrelation()
Dim MainWS As Worksheet: Set MainWS = Sheets("Time Series")
MainWS.Activate
If Cells(1, 1).Value <> "Date" Then
    MsgBox "Only Dates allowed in column A of the Time Series Worksheet", vbOKOnly + vbDefaultButton2, "Incorrect format found"
    Exit Sub
End If

'Get all entries from time series as range
Dim FirstRow As Long, FirstColumn As Long, LastRow As Long, LastColumn As Long
FirstRow = Cells.Find(what:="*", searchdirection:=xlNext, searchorder:=xlByRows).Row
FirstColumn = Cells.Find(what:="*", searchdirection:=xlNext, searchorder:=xlByColumns).Column
LastRow = Cells.Find(what:="*", searchdirection:=xlPrevious, searchorder:=xlByRows).Row
LastColumn = Cells.Find(what:="*", searchdirection:=xlPrevious, searchorder:=xlByColumns).Column
Dim rngTimeSeriesTable As Range: Set rngTimeSeriesTable = MainWS.Range(Cells(FirstRow, "B"), Cells(LastRow, LastColumn))
Dim X As Variant
Dim r As Long, c As Long
X = rngTimeSeriesTable.Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'VOLATILITY TABLE
'Loop through the variant array
For r = 2 To UBound(X, 1)
If r = LastRow Then GoTo UnloadArray
    For c = 1 To UBound(X, 2)
        'Divide
        X(r, c) = X(r, c) / X(r + 1, c)
        Debug.Print X(r, c)
    Next c
Next r
UnloadArray:
rngTimeSeriesTable = X
rngTimeSeriesTable.NumberFormat = "0.00%"
''''''''''''''''''''''''''''''''''''''''''''''''''''

'Make Ranges And Name
Dim CC As Long
Dim strAsset As String
Dim N As Name
For CC = 2 To LastColumn
    strAsset = ReplaceIllegalChars(Cells(1, CC).Value)
    For Each N In ActiveWorkbook.Names
        Debug.Print "Checked for Name: " & N.NameLocal
        If N = strAsset Then ActiveWorkbook.Names(strAsset).Delete
    Next N
    Cells(1, CC).Offset(1, 0).Select
    Range(ActiveCell, ActiveCell.End(xlDown)).Select
    Debug.Print Cells(1, CC).Value
    On Error GoTo 0
    ActiveWorkbook.Names.Add Name:=strAsset, RefersTo:=Selection
Next CC
End Sub
Function ReplaceIllegalChars(strInput As String) As String
  Dim illegal As Variant
  Dim i As Integer
    
  illegal = Array("~", "!", "?", "<", ">", "[", "]", ":", "|", _
        "*", "/", " ")
    
  For i = LBound(illegal) To UBound(illegal)
      Do While InStr(strInput, illegal(i))
          Mid(strInput, InStr(strInput, illegal(i)), 1) = "_"
      Loop
  Next i
    
  ReplaceIllegalChars = strInput
End Function

