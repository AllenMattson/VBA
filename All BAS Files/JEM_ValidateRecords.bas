Attribute VB_Name = "JEM_ValidateRecords"
Option Explicit
Sub ProcessBatchRow(bu As String, act As String, dept As String, prod As String, proj As String)
On Error GoTo Err_Handle
Dim strConn As String
Dim Select_Business_Unit As String: Select_Business_Unit = "SELECT Code as Business_Unit FROM [Hubbard Broadcasting Inc_$Dimension Value] WHERE [Dimension Code] = 'BU' AND blocked = 0"
Dim Select_Account  As String: Select_Account = "SELECT No_ as Account FROM [Hubbard Broadcasting Inc_$G_L Account] WHERE [No_]='ACT' AND blocked = 0"
Dim Select_Dept  As String: Select_Dept = "SELECT Code as deptid FROM [Hubbard Broadcasting Inc_$Dimension Value] WHERE [Dimension Code] = 'DEPT' AND blocked = 0"
Dim Select_Product  As String: Select_Product = "SELECT Code as product FROM [Hubbard Broadcasting Inc_$Dimension Value] WHERE [Dimension Code] = 'PROD' AND blocked = 0"
Dim Select_Project  As String: Select_Project = "SELECT Code as project_id FROM [Hubbard Broadcasting Inc_$Dimension Value] WHERE [Dimension Code] = 'PROJ' AND blocked = 0"

Dim Conn As ADODB.Connection
Set Conn = New ADODB.Connection
strConn = ADOconn
Conn.ConnectionString = strConn

Select_Business_Unit = Replace(Select_Business_Unit, "'BU'", "'" & bu & "'")
Select_Account = Replace(Select_Account, "'ACT'", "'" & act & "'")
Select_Dept = Replace(Select_Dept, "'DEPT'", "'" & dept & "'")
Select_Product = Replace(Select_Product, "'PROD'", "'" & prod & "'")
Select_Project = Replace(Select_Project, "'PROJ'", "'" & proj & "'")

Dim r1 As ADODB.Recordset: Set r1 = New ADODB.Recordset
Dim r2 As ADODB.Recordset: Set r2 = New ADODB.Recordset
Dim r3 As ADODB.Recordset: Set r3 = New ADODB.Recordset
Dim r4 As ADODB.Recordset: Set r4 = New ADODB.Recordset
Dim r5 As ADODB.Recordset: Set r5 = New ADODB.Recordset
If Len(bu) = 0 Then GoTo Rec2
r1.Open Select_Business_Unit, Conn, adOpenKeyset, adLockReadOnly
If HaveRecords(r1) Then boBU = True
Rec2:
If Len(act) = 0 Then GoTo Rec3
r2.Open Select_Account, Conn, adOpenKeyset, adLockReadOnly
If HaveRecords(r2) Then boACT = True
Rec3:
If Len(dept) = 0 Then GoTo Rec4
r3.Open Select_Dept, Conn, adOpenKeyset, adLockReadOnly
If HaveRecords(r3) Then boDEPT = True
Rec4:
If Len(prod) = 0 Then GoTo Rec5
r4.Open Select_Product, Conn, adOpenKeyset, adLockReadOnly
If HaveRecords(r4) Then boPROD = True
Rec5:
If Len(proj) = 0 Then GoTo ExitQuery
r5.Open Select_Project, Conn, adOpenKeyset, adLockReadOnly
If HaveRecords(r5) Then boPROJ = True
'Close Connection
ExitQuery:
On Error Resume Next
r1.Close
Set r1 = Nothing
r2.Close
Set r2 = Nothing
r3.Close
Set r3 = Nothing
r4.Close
Set r4 = Nothing
r5.Close
Set r5 = Nothing

Conn.Close
Set Conn = Nothing
Exit Sub

'Alert user about error
Err_Handle:
'MsgBox Err.Number & vbNewLine & "Description: " & Err.Description & vbNewLine & "On Code Line: " & Erl
Debug.Print Err.Number & vbNewLine & "Description: " & Err.Description & vbNewLine & Erl
Resume ExitQuery
End Sub
Function LogQuery(ByVal query As String)
ActiveSheet.Cells(row, "M").value = query
End Function
Public Function HaveRecords(rstData As ADODB.Recordset, Optional ByRef lngRecordCount As Long) As Boolean
    On Error GoTo HaveRecords_ERR
    HaveRecords = False
    If Not rstData Is Nothing Then
        If rstData.State = adStateOpen Then
            If (Not rstData.EOF) And (Not rstData.BOF) Then
                rstData.MoveFirst
                lngRecordCount = rstData.RecordCount
                Debug.Print lngRecordCount
                HaveRecords = True
            End If
        Else
            HaveRecords = False
        End If
    Else
        HaveRecords = False
    End If
    Exit Function
HaveRecords_ERR:
    MsgBox "[" & Err.Number & "] " & Err.Description, vbInformation, "HaveRecords - Error"
End Function
Sub ADO_NavisionOpenDatabase()
On Error GoTo Err_Handle
Dim Batch_Name As String
Dim strSQL As String
Dim Conn As ADODB.Connection
Dim rst As ADODB.Recordset
Dim FilterRst As ADODB.Recordset
Set Conn = New ADODB.Connection
Conn.Open ADOconn
Set rst = New ADODB.Recordset
strSQL = "SELECT Name FROM [Hubbard Broadcasting Inc_$Gen_ Journal Batch] WHERE [Journal Template Name]='GENERAL' AND [NAME]='?NAME?'"
Batch_Name = "JEM2018" 'Range("E3").value
strSQL = Replace(strSQL, "?NAME?", Batch_Name)
'LogQuery (strSQL)
'Debug.Print strSQL
Dim strr As String

rst.Open strSQL, ADOconn, adOpenKeyset, adLockOptimistic
Debug.Print "Selected: " & rst.RecordCount & " records."
rst.MoveFirst
strr = rst.GetString
Debug.Print strr
ExitQuery:
Set rst = Nothing
Conn.Close
Set Conn = Nothing
Exit Sub
Err_Handle:
Debug.Print Err.Number & vbNewLine & "Description: " & Err.Description & vbNewLine & Erl
Resume ExitQuery
End Sub
'/* TO VALIDATE THE HEADER */

'SELECT ISNULL((SELECT CODE FROM [dbo].[Hubbard Broadcasting Inc_$Dimension Value]
'WHERE CODE = '99' AND [dimension CODE] = 'BU'),0) as BU,
'ISNULL((SELECT [Name] FROM [dbo].[Hubbard Broadcasting Inc_$Gen_ Journal Batch]
'WHERE  [Journal Template Name] = 'GENERAL' AND [Name] = 'JEM999'),0) as BATCH



'/* TO VALIDATE EACH LINE */

'SELECT ISNULL((SELECT CODE FROM [dbo].[Hubbard Broadcasting Inc_$Dimension Value]
'WHERE CODE = '99' AND [dimension CODE] = 'BU'),0) as BU,
'ISNULL((SELECT CODE FROM [dbo].[Hubbard Broadcasting Inc_$Dimension Value]
'WHERE CODE = '99' AND [dimension CODE] = 'DEPT'),0) as DEPT,
'ISNULL((SELECT CODE FROM [dbo].[Hubbard Broadcasting Inc_$Dimension Value]
'WHERE CODE = '99' AND [dimension CODE] = 'PROD'),0) as PROD,
'ISNULL((SELECT CODE FROM [dbo].[Hubbard Broadcasting Inc_$Dimension Value]
'WHERE CODE = '99' AND [dimension CODE] = 'PROJ'),0) as PROJ,
'ISNULL((SELECT [No_] FROM [dbo].[Hubbard Broadcasting Inc_$G_L Account]
'WHERE [No_] = '99'),0) as ACCT


Sub ValidateTableLine(bu As String, act As String, dept As String, _
prod As String, proj As String)
Dim Batch_Num As String
Batch_Num = Range("E3").value
Dim fld As Field
Dim rst As ADODB.Recordset
Dim Conn As ADODB.Connection
Set Conn = New ADODB.Connection
Conn.Open ADOconn
boBU = False
boACT = False
boDEPT = False
boPROD = False
boPROJ = False
Dim strSQL As String
'QUERY COMMAND TO VALIDATE EACH LINE
strSQL = _
    "SELECT ISNULL((SELECT CODE FROM [dbo].[Hubbard BroadcASting Inc_$Dimension Value]"
    
strSQL = strSQL & " " & _
    "WHERE CODE ='" & Batch_Num & "' AND [dimension CODE] = '" & bu & "'),0) AS BU, " & _
    "ISNULL((SELECT CODE FROM [dbo].[Hubbard BroadcASting Inc_$Dimension Value]"

strSQL = strSQL & " " & _
    "WHERE CODE ='" & Batch_Num & "' AND [dimension CODE] = '" & dept & "'),0) AS DEPT, " & _
    "ISNULL((SELECT CODE FROM [dbo].[Hubbard BroadcASting Inc_$Dimension Value]"

strSQL = strSQL & " " & _
    "WHERE CODE ='" & Batch_Num & "' AND [dimension CODE] = '" & prod & "'),0) AS PROD, " & _
    "ISNULL((SELECT CODE FROM [dbo].[Hubbard BroadcASting Inc_$Dimension Value]"
strSQL = strSQL & " " & _
    "WHERE CODE ='" & Batch_Num & "' AND [dimension CODE] = '" & proj & "'),0) AS PROJ, " & _
    "ISNULL((SELECT [No_] FROM [dbo].[Hubbard BroadcASting Inc_$G_L Account]"
strSQL = strSQL & " " & "WHERE [No_] ='" & Batch_Num & "'),0) AS ACCT"
   
'Execute Query
Set rst = Conn.Execute(strSQL)
Debug.Print "rst string: " & rst.GetString
   
rst.Close
Set rst = Nothing
Conn.Close
Set Conn = Nothing

End Sub
Function GetRecords(db As database, ByVal query As String) As Recordset
    Dim rs As Recordset
    
    LogQuery (query)
    
    If Not db Is Nothing Then
        Set rs = db.OpenRecordset(query, dbOpenSnapshot, dbSQLPassThrough)
    Else
        Set rs = Nothing
    End If
    Set QueryDb = rs
End Function

Function IsValidBU(bu As String, rs As Recordset, Total_Records As Integer) As Boolean
    Dim found As Boolean, value As String, z As Integer
    rs.MoveFirst
    found = False
    'rs_MoveFirst rs
    
    If Not IsEmptyString(bu) Then
        For z = 0 To Total_Records - 1
            value = rs.Fields("Business_Unit").value
            If value = Format(bu, "00") Then
                found = True
                z = Total_Records - 1
            End If
            rs.MoveNext
        Next z
    Else
        found = True
    End If
    
    IsValidBU = found

End Function

Function IsValidAccount(act As String, rs As Recordset, Total_Records As Integer) As Boolean
    Dim found As Boolean, value As String, z As Integer
    rs.MoveFirst
    found = False
    'rs_MoveFirst rs
    
    If Not IsEmptyString(act) Then
        For z = 0 To Total_Records - 1
            value = rs.Fields("account").value
            If value = act Then
                found = True
                z = Total_Records - 1
            End If
            rs.MoveNext
        Next z
    Else
        found = True
    End If
    
    IsValidAccount = found

End Function

Function IsValidDept(dept As String, rs As Recordset, Total_Records As Integer) As Boolean
    Dim found As Boolean, value As String, z As Integer
    rs.MoveFirst
    found = False
    'rs_MoveFirst rs
    
    If Not IsEmptyString(dept) Then
        For z = 0 To Total_Records - 1
            value = rs.Fields("deptid").value
            If value = dept Then
                found = True
                z = Total_Records - 1
            End If
            rs.MoveNext
        Next z
    Else
        found = True
    End If
    
    IsValidDept = found

End Function

Function IsValidProd(prod As String, rs As Recordset, Total_Records As Integer) As Boolean
    Dim found As Boolean, value As String, z As Integer
    rs.MoveFirst
    found = False
    'rs_MoveFirst rs
    
    If Not IsEmptyString(prod) Then
        For z = 0 To Total_Records - 1
            value = rs.Fields("product").value
            If value = CStr(prod) Then
                found = True
                z = Total_Records - 1
            End If
            rs.MoveNext
        Next z
    Else: found = True
    End If
    
    IsValidProd = found

End Function

Function IsValidProj(proj As String, rs As Recordset, Total_Records As Integer) As Boolean
    Dim found As Boolean, value As String, z As Integer
    rs.MoveFirst
    found = False
    'rs_MoveFirst rs
    
    If Not IsEmptyString(proj) Then
        For z = 0 To Total_Records - 1
            value = rs.Fields("project_id").value
            If value = CStr(proj) Then
                found = True
                z = Total_Records - 1
            End If
            rs.MoveNext
        Next z
    Else: found = True
    End If
    
    IsValidProj = found

End Function


