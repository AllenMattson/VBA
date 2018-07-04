Attribute VB_Name = "Rst_Disconnected"
'C:\Users\Allen\Documents\Acc16Files\VBAAccess2016_ByExample\Northwind.mdb
Option Explicit

Sub Rst_Disconnected()
  Dim conn As ADODB.Connection
  Dim rst As ADODB.Recordset
  Dim strConn As String
  Dim strSQL As String
  Dim strRst As String
  Dim strFilePath As String
  
  strFilePath = "C:\Users\Allen\Documents\Acc16Files\VBAAccess2016_ByExample\Northwind.mdb"
  strSQL = "SELECT * FROM Orders WHERE CustomerID = 'VINET'"
  
  strConn = "Provider=Microsoft.Jet.OLEDB.4.0;"
  strConn = strConn & "Data Source = " & strFilePath
  Set conn = New ADODB.Connection
  conn.ConnectionString = strConn
  conn.Open
  
  Set rst = New ADODB.Recordset
  Set rst.ActiveConnection = conn
  
  ' retrieve the data
  rst.CursorLocation = adUseClient
  rst.LockType = adLockBatchOptimistic
  rst.CursorType = adOpenStatic
  rst.Open strSQL, , , , adCmdText
  
  ' disconnect the recordset
  Set rst.ActiveConnection = Nothing
   Dim f As Field
   Dim LFf As Long: LFf = rst.Fields.Count
   Dim arr() As String
   ReDim arr(0 To LFf)
   Dim i As Long
   Dim a As String
   Debug.Print "Fields: "
For Each f In rst.Fields
    arr(i) = f.Name
    a = a & arr(i) & ","
   ' Debug.Print f.Name
Next f
Debug.Print a
Debug.Print rst.GetString
Debug.Print "Total Records returned: " & rst.RecordCount
Debug.Print "----------------------------"
  ' change the CustomerID in the first record to 'OCEAN'
  rst.MoveFirst
 ' Debug.Print rst.GetString(adClipString, , ",")
  Debug.Print rst.Fields(0) & " was previously: " _
    & rst.Fields(1)
  rst.Fields("CustomerID").Value = "OCEAN"
  rst.Update

  ' stream out the recordset as a comma-delimited string
  strRst = rst.GetString(adClipString, , ",")
  Debug.Print strRst
End Sub

Sub zzzRst_Disconnected()
Dim conn As ADODB.Connection
Dim rst As ADODB.Connection
Dim strRst As String
Dim strConn As String
Dim strSQL As String
Dim strFilePath As String
strFilePath = "C:\Users\Allen\Documents\Acc16Files\VBAAccess2016_ByExample\Northwind.mdb"
strSQL = "SELECT * FROM Orders WHERE CustomerID= 'VINET'"
'strCONN = "Provider=Microsoft.jet.oledb.4.0"
strConn = "Provider='SQLOLEDB';" & _
            "Data Source='" & strFilePath & "';" & _
            "Initial Catalog='Northwind';" & _
            "Integrated Security='SSPI';" & _
            "Initial catalog=Northwind"
'strCONN = strCONN & "Data Source = " & strFilePath
Set conn = New ADODB.Connection
conn.ConnectionString = strConn
conn.Open

Set rst = New ADODB.Connection
Set rst.ActiveConnection = conn
'retrieve data
rst.CursorLocation = adUseClient
rst.LockType = adLockBatchOptimistic
rst.CursorType = adOpenStatic
rst.Open strSQL, , , adCmdText
'disconnect the recordset
Set rst.ActiveConnection = Nothing
'stream out recordset
strRst = rst.GetString(adClipString, , ",")
Debug.Print strRst

End Sub
