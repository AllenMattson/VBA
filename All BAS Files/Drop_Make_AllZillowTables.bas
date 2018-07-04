Attribute VB_Name = "Drop_Make_AllZillowTables"
Private FileRows As Long
Sub ATTEMPT()
Sheets("Sheet1").Activate
Dim cell As Range
Dim MyRNG As Range: Set MyRNG = Range("B3:B94")
Dim MyFile_Name As String
Dim SQLCommand As String
Dim SQLCommand2 As String
For Each cell In MyRNG
    If Left(cell.Value, 10) = "Zip_Median" Then
        Drop_Make_AllZillowTables.ADOReadCSVFile (cell.Value)
        MyFile_Name = cell.Value
        MyFile_Name = Left(MyFile_Name, Len(MyFile_Name) - 4)
        'Debug.Print "--" & MyFile_Name & " Table is dropped and created here"
        'SQLCommand = "--" & MyFile_Name & " Table is dropped and created here"
        SQLCommand = vbNewLine & " DROP TABLE IF EXISTS GeoCityDB.dbo." & MyFile_Name & ";" & vbNewLine
        SQLCommand = SQLCommand & " CREATE TABLE GeoCityDB.dbo." & MyFile_Name
        SQLCommand = SQLCommand & " (MonthDate VARCHAR(255) NULL,ZipCode VARCHAR(255) NULL," & MyFile_Name
        SQLCommand = SQLCommand & " VARCHAR(255) NULL)" & vbNewLine & " ON [PRIMARY]"
        SQLCommand = SQLCommand & " BULK INSERT GeoCityDB.dbo." & MyFile_Name & vbNewLine
        SQLCommand = SQLCommand & " FROM 'C:\Users\Allen\Documents\Visual Studio 2015\Projects\GeoDBFile\data\" & MyFile_Name & ".csv'"
        SQLCommand = SQLCommand & " WITH(FIRSTROW=1,FIELDTERMINATOR = ',',LASTROW=" & FileRows & " ,ROWTERMINATOR = '0x0a');"
        'Debug.Print SQLCommand
        'Debug.Print
        SQLCommand2 = Trim(SQLCommand2) & vbNewLine & Trim(SQLCommand)
    End If
Next cell
SQLCommand2 = Trim(SQLCommand2)
Debug.Print SQLCommand2
Sheets.Add: Range("A1:G3").Cells.Merge: Range("A1").Value = SQLCommand2
End Sub
Sub ADOReadCSVFile(CSV_FILE As String)
'Const CSV_FILE As String = "TESTFILE.txt"
Dim strPathtoTextFile As String
'set reference to Microsoft ActiveX Data Objects Library (Tools>References...)
Dim objConnection As ADODB.Connection
Dim objRecordset As ADODB.Recordset

Set objConnection = New ADODB.Connection
Set objRecordset = New ADODB.Recordset
'strPathtoTextFile = ThisWorkbook.Path & "\"

objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
"Data Source=C:\Users\Allen\Documents\Visual Studio 2015\Projects\GeoDBFile\data\backup\;" & _
"Extended Properties=""text;HDR=No;FMT=CSVDelimited"""

objRecordset.Open "SELECT * FROM " & CSV_FILE, objConnection, adOpenStatic, adLockOptimistic, adCmdText

If Not objRecordset.EOF Then FileRows = objRecordset.RecordCount 'Debug.Print CSV_FILE & ": " & objRecordset.RecordCount
FileRows = FileRows - 1
objRecordset.Close 'Close ADO objects
objConnection.Close

Set objRecordset = Nothing
Set objConnection = Nothing
End Sub

