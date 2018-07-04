Attribute VB_Name = "JEM_ValidateEntries"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MODULE: JEM_ValidateEntries
'AUTHOR: ALLEN MATTSON
'DATE: 5/5/2018
'===============================================================================
'DESCRIPTION:
'   VALIDATE THE USER INPUTS AND CHECK FOR INVALID DATABASE ENTRIES
'   THE FIRST PART WILL CHECK FOR VALID INPUTS THAT ARE MANDATORY TO RUN
'   THE TABLE QUERY, IF THE VALUE ISN'T VALID, ALERT USER, VALUES TURN RED AND EXIT SUB
'===============================================================================
'   THE SECOND PORTION IDENTIFIES TABLE ENTRIES FROM RANGE("A6") DOWN
'   EACH ROW OF THE TABLE MUST RETURN AT LEAST 1 VALUE MATCH IN PER VARIABLE
'   THE MACRO TO PASS THE VARIABLES IS USED WHERE THE CODE LINE IS:
'   --->  Call JEM_ValidateEntries.QueryEachTableLine(bu, dept, prod, proj, act) ***********
'   COLUMNS D,E,F,G,H ----->FOR EACH ROW PASS THE VALUES IN THOSE COLUMNS
'   AFTER THE TABLE HAS LOOPED, ANY RECORDS THAT ARE INVALID ARE DISPLAYED IN RED
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
'===========CONNECTION STRING========================
Public Const ADOconn = "Driver=SQL SERVER; Trusted_Connection=no; Server=hbi-nav-sql; Database=HBI; Uid=jem; Pwd=Jem2013!"
'====================================================
Public TotalRecs As Long
Public boBU As Boolean
Public boACT As Boolean
Public boDEPT As Boolean
Public boPROD As Boolean
Public boPROJ As Boolean
Sub ValidateEntries()
'QUIT IF THIS ISN'T THE RIGHT WORKSHEET
If Range("A5").Value <> "Description" Then Exit Sub
'Get the batch_name at the top and is used when passed into the
'   ValidBatch function in the ValidateRecords Module.
'   Alert user and exit sub if function fails.

Dim bu As String
Dim act As String
Dim dept As String
Dim prod As String
Dim proj As String
Dim ThisDesc As String
Dim PrevDesc As String
Dim Batch_Name As String
Dim IsValidRow As Boolean
'Check for date
If Not IsDate(ActiveSheet.Range("A3")) Then
    Range("A3").Font.Color = vbRed
    MsgBox ("Enter Date in cell A3")
    Exit Sub
Else
    Range("A3").Font.Color = vbBlack
End If
'Check Journal number
If Len(Trim(Range("J3"))) = 0 Or Len(Trim(Range("J3").Value)) > 8 Or IsNull(Range("J3")) Then
    Range("J3").Font.Color = vbRed
    MsgBox ("Enter valid journal number in cell J3")
    Exit Sub
Else
    Range("J3").Value = vbBlack
End If
'Check Business Unit Name
If Not IsNumeric(Range("I3")) Then
    Range("I3").Font.Color = vbRed
    MsgBox ("Enter a numeric BU# in cell I3")
    Exit Sub
End If
''''''''''''Function to test if valid batch''''''''''''''
If JEM_ValidateEntries.ValidBatch(Range("E3").Value) <> True Then Exit Sub
Dim iCheck As Long
Dim iLR As Long
'Find last row in entries by keying from acct# current region
With ActiveSheet.Range("H6").CurrentRegion
    iLR = .Rows(.Rows.count).row
    If iLR < 7 Then
        MsgBox "Not enough entries found"
        Exit Sub
    End If
End With
'Double Check Last Row of Table in case any lines were missed 5/15/2018
If Len(Range("A1000")) > 0 Then
    Dim Dlr As Long, Clr As Long
    Rows("1000:1000").EntireRow.Hidden = True
    Dlr = Cells(Rows.count, "I").End(xlUp).row
    Clr = Cells(Rows.count, "J").End(xlUp).row
    If Dlr > iLR Then iLR = Dlr
    If Clr > iLR Then iLR = Clr
    Rows("1000:1000").EntireRow.Hidden = False
End If
ActiveSheet.Range("K6:K1000").ClearContents
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'LOOP THE TABLE AND IF VALUES AREN'T FORMATTED CORRECTLY, MAKE VALUES RED
'IDENTIFY BU, PROD,PROJ,DEPT,ACT VALUES AND CHECK IF FOUND
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
For iCheck = 7 To iLR
IsValidRow = True 'Set isValidRow to true, if row fails any validation checks, set isvalidrow=false
ThisDesc = Cells(iCheck, "A")
    If Len(Trim(Cells(iCheck, "A").Value)) = 0 Then
        If Len(PrevDesc) = 0 Then
            Cells(iCheck, "A").Font.Color = vbRed
        Else
            Cells(iCheck, "A").Font.Color = vbBlack
        End If
    Else
        PrevDesc = ThisDesc
            If Len(ThisDesc) > 50 Then 'Description must be under 50 characters
                Cells(iCheck, "A").Font.Color = vbRed
                IsValidRow = False
            Else
                Cells(iCheck, "A").Font.Color = vbBlack
            End If
    End If

    'CHECK NUMBER FORMATTING IN ROW
    If Not IsNumeric(Cells(iCheck, "I")) Then
        Cells(iCheck, "I").Font.Color = vbRed
        IsValidRow = False
    End If
    If IsNumeric(Cells(iCheck, "I")) Then Cells(iCheck, "I").Font.Color = vbBlack
    If Not IsNumeric(Cells(iCheck, "J")) Then
        Cells(iCheck, "J").Font.Color = vbRed
        IsValidRow = False
    End If
    If IsNumeric(Cells(iCheck, "J")) Then Cells(iCheck, "J").Font.Color = vbBlack
    
    'IF ROW ISN'T BLANK BUT BU IS MISSING FROM COLUMN F, THEN BU IS THE VALUE FROM I3
    'IF THIS IS A BLANK ROW, GOTO NEXT iCheck
    If Len(Trim(ActiveSheet.Cells(iCheck, "F"))) = 0 And _
            (Len(Trim(Cells(iCheck, "I"))) <> 0 Or _
            Len(Trim(Cells(iCheck, "J"))) <> 0) Then
            ActiveSheet.Cells(iCheck, "F").Value = Cells(3, "I").Value
    Else
            If Len(Trim(ActiveSheet.Cells(iCheck, "F"))) = 0 And _
            (Len(Trim(Cells(iCheck, "I"))) = 0 And _
            Len(Trim(Cells(iCheck, "J"))) = 0) Then GoTo NextiCheck
    End If
'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'CALL SUB WHICH VALIDATES AGAINST DB AND FORMATS INVALID/VALID ENTRIES OF VARIABLES
    bu = Trim(Cells(iCheck, "F"))
    act = Trim(Cells(iCheck, "H"))
    prod = Trim(Cells(iCheck, "D"))
    proj = Trim(Cells(iCheck, "E"))
    dept = Trim(Cells(iCheck, "G"))
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call JEM_ValidateEntries.QueryEachTableLine(bu, dept, prod, proj, act)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^'
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'PUBLIC APPLY FORMATTING - BY CHECKING BOOLEAN RESPONSES
    'bo is prefixed to variables representing it as a boolean expression
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    With ActiveSheet
        If Not boBU And Len(bu) > 0 Then
            Cells(iCheck, "F").Font.Color = vbRed
            IsValidRow = False
        Else
            Cells(iCheck, "F").Font.Color = vbBlack
        End If
        If Not boACT And Len(act) > 0 Then
            Cells(iCheck, "H").Font.Color = vbRed
            IsValidRow = False
        Else
            Cells(iCheck, "H").Font.Color = vbBlack
        End If
        If Not boPROD And Len(prod) > 0 Then
            Cells(iCheck, "D").Font.Color = vbRed
            IsValidRow = False
        Else
            Cells(iCheck, "D").Font.Color = vbBlack
        End If
        If Not boPROJ And Len(proj) > 0 Then
            Cells(iCheck, "E").Font.Color = vbRed
            IsValidRow = False
        Else
            Cells(iCheck, "E").Font.Color = vbBlack
        End If
        If Not boDEPT And Len(dept) > 0 Then
            Cells(iCheck, "G").Font.Color = vbRed
            IsValidRow = False
        Else
            Cells(iCheck, "G").Font.Color = vbBlack
        End If
        If IsValidRow Then
            With Cells(iCheck, "K") 'Checkmark
                .Font.Name = "Wingdings"
                .Value = "ü"
            End With
        End If
End With
NextiCheck:
Next iCheck
End Sub
Public Function ValidBatch(Batch_Name As String) As Boolean

'Open Connection, Check batch and return true or false if any records are found
On Error GoTo Err_Handle

Dim strSQL As String
Dim Conn As adodb.Connection
Dim rst As adodb.Recordset
Set Conn = New adodb.Connection
Set rst = New adodb.Recordset
Conn.Open ADOconn
strSQL = "SELECT Name FROM [Hubbard Broadcasting Inc_$Gen_ Journal Batch] WHERE [Journal Template Name]='GENERAL' AND [NAME]='?NAME?'"
strSQL = Replace(strSQL, "?NAME?", Batch_Name)
rst.Open strSQL, ADOconn, adOpenKeyset, adLockOptimistic
'Debug.Print "Selected: " & rst.RecordCount & " records."
If rst.RecordCount > 0 Then
    ValidBatch = True
    Range("E3").Font.Color = vbBlack
Else
    ValidBatch = False
    Range("E3").Font.Color = vbRed
    MsgBox ("Enter a valid batch number in cell E3")
    Exit Function
End If
'Close Connection
ExitQuery:
On Error Resume Next
rst.Close
Set rst = Nothing
Conn.Close
Set Conn = Nothing
Exit Function

'Alert user about error
Err_Handle:
MsgBox Err.Number & vbNewLine & "Description: " & Err.Description & vbNewLine & Erl
'Debug.Print Err.Number & vbNewLine & "Description: " & Err.Description & vbNewLine & Erl
Resume ExitQuery
End Function
Sub QueryEachTableLine(bu As String, dept As String, prod As String, _
proj As String, act As String)
Dim fld As Field
Dim rst As adodb.Recordset
Dim Conn As adodb.Connection
'/* TO VALIDATE THE HEADER */
Dim Batch_Num As String: Batch_Num = Range("E3").Value
'SELECT ISNULL((SELECT CODE FROM [dbo].[Hubbard Broadcasting Inc_$Dimension Value]
'WHERE CODE = '99' AND [dimension CODE] = 'BU'),0) as BU,
'ISNULL((SELECT [Name] FROM [dbo].[Hubbard Broadcasting Inc_$Gen_ Journal Batch]
'WHERE  [Journal Template Name] = 'GENERAL' AND [Name] = 'JEM999'),0) as BATCH

'/* TO VALIDATE EACH LINE */
Dim strQuery As String
strQuery = "SELECT ISNULL((SELECT CODE FROM DBO.[Hubbard Broadcasting Inc_$Dimension Value] " & _
    "WHERE CODE='" & bu & "' AND [dimension CODE]='BU'),0) as BU, " & _
    "ISNULL((SELECT CODE FROM DBO.[Hubbard Broadcasting Inc_$Dimension Value] " & _
    "WHERE CODE='" & dept & "' AND [dimension CODE] = 'DEPT'),0) as DEPT, " & _
    "ISNULL((SELECT CODE FROM [dbo].[Hubbard Broadcasting Inc_$Dimension Value] " & _
    "WHERE CODE ='" & prod & "' AND [dimension CODE] = 'PROD'),0) as PROD, " & _
    "ISNULL((SELECT CODE FROM [dbo].[Hubbard Broadcasting Inc_$Dimension Value]" & _
    "WHERE CODE ='" & proj & "' AND [dimension CODE] = 'PROJ'),0) as PROJ, " & _
    "ISNULL((SELECT [No_] FROM [dbo].[Hubbard Broadcasting Inc_$G_L Account] " & _
    "WHERE [No_] = '" & act & "'),0) as ACCT"
Set Conn = New adodb.Connection
Conn.Open ADOconn
Dim strCONCAT As String
Set rst = Conn.Execute(strQuery)

rst.MoveFirst
   Do Until rst.EOF
    'Set unfound records to false
        If rst.Fields("bu").Value > 0 Then
            boBU = True
        Else
            boBU = False
        End If
        If rst.Fields("dept").Value > 0 Then
            boDEPT = True
        Else
            boDEPT = False
        End If
        If rst.Fields("prod").Value > 0 Then
            boPROD = True
        Else
            boPROD = False
        End If
        If rst.Fields("proj").Value > 0 Then
            boPROJ = True
        Else
            boPROJ = False
        End If
        If rst.Fields("acct").Value > 0 Then
            boACT = True
        Else
            boACT = False
        End If
       rst.MoveNext
   Loop
 End Sub




