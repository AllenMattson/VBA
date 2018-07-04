Attribute VB_Name = "Append2CSV"
Sub Append2CSV()
Dim tmpCSV As String 'string to hold the CSV info
Dim f As Integer
Dim The_Name As String: The_Name = "test"
Dim CSVFile As String: CSVFile = "C:\Users\Allen\Documents\Visual Studio 2015\Projects\GeoDBFile\data\" 'replace with your filename
CSVFile = CSVFile & The_Name & ".csv"
f = FreeFile

Open CSVFile For Append As #f
tmpCSV = Range2CSV(Selection)
Print #f, tmpCSV
Close #f
End Sub

Function Range2CSV(list) As String
Dim tmp As String
Dim cr As Long
Dim r As Range

If TypeName(list) = "Range" Then
cr = 1

For Each r In list.Cells
If r.Row = cr Then
If tmp = vbNullString Then
tmp = r.Value
Else
tmp = tmp & "," & r.Value
End If
Else
cr = cr + 1
If tmp = vbNullString Then
tmp = r.Value
Else
tmp = tmp & Chr(10) & r.Value
End If
End If
Next
End If

Range2CSV = tmp
End Function
