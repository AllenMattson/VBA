Attribute VB_Name = "Build_ZillowMasterData"
'DESCRIPTION:
'Loop through MS SQL database
'Sort values into date oriented columns
'insert into new database table
'''''''''''''''''''''''''''''''''''''''''
Sub LoopZips()
Sheets("Zip").Activate
Dim MyZip As String
Dim i As Integer, lr As Integer: lr = Cells(Rows.Count, 1).End(xlUp).Row
For i = lr To lr - 10 Step -1
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'NEED LOOP TO STOP WHEN i=800 AND MOVE DATA TO DATABASE
    'CLEAR THE TEST SHEET
    'LEAVE ROW 1 OF TEST SHEET
    'THEN ASSIGN MYZIP WITH THE VALUE OF i
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    MyZip = Sheets("Zip").Cells(i, 1).Value
    If MyZip = "" Then GoTo NewZip
    ConnectSqlServer (MyZip)
    SeperateData
NewZip:
Next i
End Sub
Sub ConnectSqlServer(MyZip As String)
Sheets("Data").Activate: Rows("2:1000000").EntireRow.ClearContents
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim sConnString As String
 
    ' Create the connection string.
    sConnString = "Provider=SQLOLEDB;Data Source=ALLENSDESKTOP;" & _
                  "Initial Catalog=GeoCityDB;" & _
                  "Integrated Security=SSPI;"
    
    ' Create the Connection and Recordset objects.
    Set conn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    ' Open the connection and execute.
    conn.Open sConnString
    Set rs = conn.Execute("SELECT * FROM PricePerSqFt_by_ZipCode WHERE zip=" & MyZip & ";")
    
    ' Check we have data.
    If Not rs.EOF Then
        ' Transfer result.
        Sheets("Data").Range("A2").CopyFromRecordset rs
    ' Close the recordset
        rs.Close
    Else
        MsgBox "Error: No records returned.", vbCritical
    End If

    ' Clean up
    If CBool(conn.State And adStateOpen) Then conn.Close
    Set conn = Nothing
    Set rs = Nothing
    
End Sub
Sub SeperateData()
Dim NewSh As Worksheet
Dim SH As Worksheet: Set SH = Sheets("Data")
Dim MyZip As String: MyZip = SH.Cells(2, 2).Value
Set NewSh = Sheets("Test")
'Cells(1, 1).Value = "Year": Cells(1, 2).Value = "Month": Cells(1, 3).Value = "PerSQFT": Cells(1, 3).Value = "Zip"
Dim zc_ID As Long
Dim MyYear As String
Dim MyMonth As String
Dim MyState As String
Dim MyCity As String
Dim MyMetro As String
Dim MyCounty As String
Dim MySQFT As Double

Dim MyVal As Double
SH.Activate
Dim LC As Integer: LC = Cells(1, Columns.Count).End(xlToLeft).Column
Dim i As Integer
'Loop column headers to place data in rows
For i = 8 To LC
'Constants
        zc_ID = Cells(2, 1)
        MyZip = Cells(2, 2)
        MyCity = Cells(2, 3)
        MyState = Cells(2, 4)
        MyMetro = Cells(2, 5)
        MyCounty = Cells(2, 6)

        '''''''
        'values
        '''''''
        MyYear = Mid(Cells(1, i), 9, 4)
        MyMonth = Right(Cells(1, i), 2)
        MySQFT = Cells(2, i)
        '''''''''''''
        NewSh.Cells(1000000, 1).End(xlUp).Offset(1, 0).Value = zc_ID
        NewSh.Cells(1000000, 2).End(xlUp).Offset(1, 0).Value = MyYear
        NewSh.Cells(1000000, 3).End(xlUp).Offset(1, 0).Value = MyMonth
        NewSh.Cells(1000000, 4).End(xlUp).Offset(1, 0).Value = MyState
        NewSh.Cells(1000000, 5).End(xlUp).Offset(1, 0).Value = MyCity
        NewSh.Cells(1000000, 6).End(xlUp).Offset(1, 0).Value = MyZip
        NewSh.Cells(1000000, 7).End(xlUp).Offset(1, 0).Value = MyMetro
        NewSh.Cells(1000000, 8).End(xlUp).Offset(1, 0).Value = MyCounty
        NewSh.Cells(1000000, 9).End(xlUp).Offset(1, 0).Value = MySQFT
Next i
End Sub
