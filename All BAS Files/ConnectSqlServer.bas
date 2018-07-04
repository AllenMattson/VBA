Attribute VB_Name = "ConnectSqlServer"
Sub ConnectSqlServer()
Sheets("Data").Activate: Cells.Clear
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
    Set rs = conn.Execute("SELECT * FROM IPV4;")
    
    ' Check we have data.
    If Not rs.EOF Then
        ' Transfer result.
        Sheets("Data").Range("A1").CopyFromRecordset rs
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
