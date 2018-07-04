Attribute VB_Name = "ADO_Test"
Option Explicit
Public Const BatchCMD = "SELECT Name FROM [HubbardBroadcasting Inc_$G_JournalBatch] WHERE [Journal Template Name]='GENERAL' AND [NAME]='?NAME?'"
Public Const DBBconn = "ODBC;Driver={SQL Server}; Trusted_Connection=no; Server=hbi-nav-sql; Database=HBI; Uid=jem; Pwd=Jem2013!"
Public Const ADOconn = "Driver=SQL SERVER; Trusted_Connection=no; Server=hbi-nav-sql; Database=HBI; Uid=jem; Pwd=Jem2013!"
Public Const DAOconn = "Driver=SQL SERVER; Database=HBI; Uid=jem; Pwd=Jem2013!;DSN=hbi-nav-sql" 'Driver={SQL Server}; Trusted_Connection=no; Server=hbi-nav-sql;
Public boBU As Boolean ': boBU = False
Public boACT As Boolean ': boACT = False
Public boDEPT As Boolean ': boDEPT = False
Public boPROD As Boolean ': boPROD = False
Public boPROJ As Boolean ': boPROJ = False
Private rst_BatchNum As DAO.Recordset
Public db As DAO.database
Private DBB As DAO.database
Sub ADO_NavisionOpenDatabase()
Dim vData
On Error GoTo Err_Handle
Dim Batch_Name As String
Dim strSQL As String
Dim Conn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim FilterRst As ADODB.Recordset
Set Conn = New ADODB.Connection
Conn.Open ADOconn 'DBBconn
Set rst = New ADODB.Recordset
'strSQL = BatchCMD
strSQL = "SELECT Name FROM [JournalBatch] WHERE [Journal Template Name]='GENERAL' AND [NAME]='?NAME?'"
Batch_Name = Range("E3").Value
strSQL = Replace(strSQL, "?NAME?", Batch_Name)
'LogQuery (strSQL)
Debug.Print strSQL

rst.Open strSQL, ADOconn, adOpenKeyset, adLockOptimistic
Debug.Print "Selected: " & rst.RecordCount & " records."
ExitQuery:
vData = rst.GetRows
rst.Close
Set rst = Nothing
Conn.Close
Set Conn = Nothing
Debug.Print vData
Set vData = Nothing
Exit Sub
Err_Handle:
Debug.Print Err.Number & vbNewLine & "Description: " & Err.Description & vbNewLine & Erl
Resume ExitQuery
End Sub
