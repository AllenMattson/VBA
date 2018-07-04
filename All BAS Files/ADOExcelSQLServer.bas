Attribute VB_Name = "ADOExcelSQLServer"
Sub ADOExcelSQLServer()

     
    Dim Cn As ADODB.Connection
    Dim Server_Name As String
    Dim Database_Name As String
    Dim User_ID As String
    Dim Password As String
    Dim SQLStr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
     
    Server_Name = "EXCEL-PC\EXCELDEVELOPER" ' Enter your server name here
    Database_Name = "AdventureWorksLT2012" ' Enter your database name here
    User_ID = "" ' enter your user ID here
    Password = "" ' Enter your password here
    SQLStr = "SELECT * FROM [SalesLT].[Customer]" ' Enter your SQL here
     
    Set Cn = New ADODB.Connection
    Cn.Open "Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & _
    ";Uid=" & User_ID & ";Pwd=" & Password & ";"
     
    rs.Open SQLStr, Cn, adOpenStatic
     ' Dump to spreadsheet
    With Worksheets("sheet1").Range("a1:z500") ' Enter your sheet name and range here
        .ClearContents
        .CopyFromRecordset rs
    End With
     '            Tidy up
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
End Sub

