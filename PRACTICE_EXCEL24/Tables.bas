Attribute VB_Name = "Tables"
Sub GetCategories_2()
    Dim conn As New ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim strPath As String
    Dim wks As Worksheet
    Dim j As Integer
    Dim rng As Range
    
    strPath = "C:\Excel2013_ByExample\Northwind.mdb"

    Set wks = ThisWorkbook.ActiveSheet

    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
              & "Data Source=" & strPath & ";"

    ' Create a Recordset from data in the Categories table

    Set rst = conn.Execute(CommandText:="Select CategoryID," & _
                "CategoryName, Description from Categories", _
                Options:=adCmdText)

    rst.MoveFirst

    ' transfer the data to Excel
    ' get the names of fields first
    With wks.Range("A1")
        .CurrentRegion.Clear
        For j = 0 To rst.Fields.Count - 1
            .Offset(0, j) = rst.Fields(j).Name
        Next j
        .Offset(1, 0).CopyFromRecordset rst
        .CurrentRegion.Columns.AutoFit
        .Cells(1, 1).Select
    End With
    rst.Close
    conn.Close

    Set rst = Nothing
    Set conn = Nothing
    
    'create a table in Excel

    Set rng = wks.Range(Range("A1").CurrentRegion.Address)
    wks.ListObjects.Add xlSrcRange, rng

End Sub

Sub GetCategories()
    Dim conn As New ADODB.Connection
    Dim rst As ADODB.Recordset
    Dim strPath As String
    Dim wks As Worksheet
    Dim j As Integer
    
    strPath = "C:\Excel2013_ByExample\Northwind.mdb"

    Set wks = ThisWorkbook.ActiveSheet

    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
              & "Data Source=" & strPath & ";"

    ' Create a Recordset from data in the Categories table

    Set rst = conn.Execute(CommandText:="Select CategoryID," & _
                "CategoryName, Description from Categories", _
                Options:=adCmdText)

    rst.MoveFirst

    ' transfer the data to Excel
    ' get the names of fields first
    With wks.Range("A1")
        .CurrentRegion.Clear
        For j = 0 To rst.Fields.Count - 1
            .Offset(0, j) = rst.Fields(j).Name
        Next j
        .Offset(1, 0).CopyFromRecordset rst
        .CurrentRegion.Columns.AutoFit
        .Cells(1, 1).Select
    End With
    rst.Close
    conn.Close

    Set rst = Nothing
    Set conn = Nothing
    
End Sub

Sub List_Headers()
    Dim rng As Range
    Dim wks As Worksheet

    Set wks = ActiveWorkbook.Worksheets(3)
    Set rng = wks.Range("A2:B5")

    wks.ListObjects.Add SourceType:=xlSrcRange, _
        Source:=rng, XlListObjectHasHeaders:=xlNo
End Sub

Sub List_Headers2()
    Dim rng As Range
    Dim wks As Worksheet

    Set wks = ActiveWorkbook.Worksheets(3)
    Set rng = wks.Range("A1:B5")

    wks.ListObjects.Add SourceType:=xlSrcRange, _
       Source:=rng, XlListObjectHasHeaders:=xlYes
End Sub

Sub DefineTableName()
        Dim wks As Worksheet
        Dim lst As ListObject

        Set wks = ActiveWorkbook.Worksheets(3)

        Set lst = wks.ListObjects(1)
        lst.Name = "1st Qtr. 2010 2013 Student Scores"
    End Sub

Sub DeleteLastCol()
    Dim myList As ListObject
    Dim lastCol As Integer

    Set myList = ActiveSheet.ListObjects(1)
    lastCol = myList.ListColumns.Count
    myList.ListColumns(lastCol).Delete
End Sub


Sub CountListRows()
    Dim objRows As ListRows
    Set objRows = ActiveSheet.ListObjects(1).ListRows
    Debug.Print objRows.Count
End Sub

Sub DefineTableName2()
    Dim wks As Worksheet
    Dim lst As ListObject
    Dim col As ListColumn
    Dim c As Variant

    Set wks = ActiveWorkbook.Worksheets(3)

    Set lst = wks.ListObjects(1)
    With lst
        .Name = "1st Qtr. 2013 Student Scores"
        .ListColumns(1).Name = "Student Name"
        .ListColumns(2).Name = "Score"
        Set col = .ListColumns.Add
        col.Name = "Previous Score"
        Debug.Print "Header Address = " & .HeaderRowRange.Address
        Debug.Print "Data Range = " & .Range.Address
        Debug.Print "Data Body Range = " & .DataBodyRange.Address

        For Each c In wks.Range(.HeaderRowRange.Address)
            Debug.Print c
        Next
    End With
End Sub






