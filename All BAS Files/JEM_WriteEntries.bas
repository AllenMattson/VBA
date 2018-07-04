Attribute VB_Name = "JEM_WriteEntries"
Option Explicit
Public Const STR_Open_Failed = "Unable to write Journal Entries.  Unable to open connection."
Dim x As String

Sub entWriteJournal()
  On Error GoTo HandleError
    
    Dim x As Long
    Dim numJEs As Long
    Dim MSG As String
    Dim query As String
    Dim rs As adodb.Recordset
    Dim PopulatedRow As String
    Dim PreviousDescription As String
    Dim MyDescription As String
    Dim MyLineNo As Long
    Dim BlankRows As Integer
    numJEs = 0
    'The following code presents a message to the user and allows for a graceful
    ' retreat if they are not prepared to write the journal entry at this time
    If Not GetUserConfirmationToWrite() Then
        Exit Sub
    End If
   
    ' Determine the next journal line number
    Navision_GetMaxLineNo (ActiveSheet.Range("E3"))
    'MyLineNo = 1 'Navision_GetMaxLineNo(ActiveSheet.Range("E3"))
    If MyLineNo > 0 Then
        ' increment max line number if one exists
        MyLineNo = MyLineNo + 1 ' 07-21-2006: use constant value instead of just incrementing by 1
    Else
        ' retrieve the first line number for this batch
        MyLineNo = Navision_GetBeginLineNo(ActiveSheet.Range("E3")) + 1 ' 07-21-2006: use constant value instead of just incrementing by 1
    End If

 Dim LASTrow As Integer
    BlankRows = 0
    For x = 7 To 999
    Debug.Print x & " : " & MyLineNo
        PopulatedRow = ActiveSheet.Cells(x, 1) & ActiveSheet.Cells(x, 2) & ActiveSheet.Cells(x, 3) & _
                       ActiveSheet.Cells(x, 4) & ActiveSheet.Cells(x, 5) & ActiveSheet.Cells(x, 6) & _
                       ActiveSheet.Cells(x, 7) & ActiveSheet.Cells(x, 8) & ActiveSheet.Cells(x, 9) & _
                       ActiveSheet.Cells(x, 10)
        If PopulatedRow = "" Then
            BlankRows = BlankRows + 1
            If BlankRows > 2 Then
                x = LASTrow
            End If
        Else
            BlankRows = 0
            'desc_col=1 or "A"
            If Not IsEmpty(ActiveSheet.Cells(x, "A")) Then
                PreviousDescription = ActiveSheet.Cells(x, "A")
            End If
            'rider_col=2 or "B"
            If Len(ActiveSheet.Cells(x, "B")) > 50 Then
                MyDescription = Left(ActiveSheet.Cells(x, "B"), 50)
            Else
                MyDescription = Left(PreviousDescription, (50 - (Len(ActiveSheet.Cells(x, "B")) + 1)))
                MyDescription = MyDescription & " " & ActiveSheet.Cells(x, "B")
            End If

    ' construct command to update DB
    '''''''''''''''
    'THIS IF STATEMENT DOESN'T LOOK RIGHT
            If Not ActiveSheet.Cells(x, 12).Value <> "" And (ActiveSheet.Cells(x, 11).Value <> "") Then
                query = _
"exec [dbo].[Insert_into_Gen_Journal_line_Nav_2013] @JournalLine = ?LINENO?, @HeaderJournalDate = '?DATE?', @HeaderBusinessUnit = '?HEADERBUSINESSUNIT?', @HeaderJournalID = '?JOURNALID?', @HeaderBatch = '?BATCH?', @LineDescription = '?DESCRIPTION?', @LineAmount = ?AMOUNT?, @LineBusinessUnit = '?LINEBUSINESSUNIT?', @LineDepartment = '?LINEDEPARTMENT?', @LineAccount = '?ACCOUNTNO?', @LineProduct = '?LINEPRODUCT?', @LineProject = '?LINEPROJECT?',@SystemDateTime=?Timestamp?"
                query = Replace(query, "?BATCH?", Range(Range("E3").Value))
                query = Replace(query, "?LINENO?", MyLineNo)
                query = Replace(query, "?ACCOUNTNO?", ActiveSheet.Cells(x, "H"))
                query = Replace(query, "?DATE?", ActiveSheet.Range("A3"))
                query = Replace(query, "?DESCRIPTION?", MyDescription)
                If Not IsEmpty(ActiveSheet.Cells(x, "J")) Then
                    query = Replace(query, "?AMOUNT?", ActiveSheet.Cells(x, "J") * -1)
                Else
                    query = Replace(query, "?AMOUNT?", ActiveSheet.Cells(x, "I"))
                End If
                query = Replace(query, "?HEADERBUSINESSUNIT?", Format(ActiveSheet.Range("I3"), "00")) ' header BU should have leading zero
                query = Replace(query, "?DEPARTMENT?", ActiveSheet.Cells(x, "G"))
                query = Replace(query, "?JOURNALID?", ActiveSheet.Range("J3"))
                If Format(ActiveSheet.Cells(x, "F")) > " " Then
                    query = Replace(query, "?LINEBUSINESSUNIT?", Format(ActiveSheet.Cells(x, "F"), "00")) ' line BU should have leading zero
                Else
                    query = Replace(query, "?LINEBUSINESSUNIT?", Format(ActiveSheet.Range("I3"), "00"))  ' line BU should have leading zero
                End If
                query = Replace(query, "?LINEDEPARTMENT?", ActiveSheet.Cells(x, "G"))
                query = Replace(query, "?LINEPRODUCT?", ActiveSheet.Cells(x, "D"))
                query = Replace(query, "?LINEPROJECT?", ActiveSheet.Cells(x, "E"))
                query = Replace(query, "?Timestamp?", "'" & Now() & "'")
                Debug.Print query
                Dim Conn As adodb.Connection
                Set Conn = New adodb.Connection
                Conn.Open ADOconn
                ' execute command
                Set rs = New adodb.Recordset
                rs.Open query, Conn, adOpenKeyset, adLockOptimistic
                Set rs = Null
                With ActiveSheet.Cells(x, 12)
                    .Font.Name = "Wingdings"
                    .Value = "ü"
                End With
                MyLineNo = MyLineNo + 1 ' 07-21-2006: use constant value instead of just incrementing by 1
                numJEs = numJEs + 1
            End If
        End If
    Next x
    
    ActiveWorkbook.Save
    MsgBox MSG, vbInformation, MSG_TITLE
    rs.Close
    Set rs = Nothing
    Exit Sub

HandleError:
      MsgBox Err.Number & ": " & Err.Description
End Sub
' Go to Gen_ Journal Line table to check if there are unposted entries for the selected BATCH.
' This function returns 0 if no line numbers were found
Function Navision_GetMaxLineNo(batch As String) As Long
Dim strSQL As String
Dim Conn As adodb.Connection
Set Conn = New adodb.Connection
Conn.Open ADOconn
strSQL = "SELECT max([Line No_]) FROM [Hubbard Broadcasting Inc_$Gen_ Journal Line] WHERE [Journal Template Name] = 'GENERAL' AND [Journal Batch Name] = '?BATCH?'"
Dim maxLineNo As Long
maxLineNo = 0
Dim rst As adodb.Recordset
Set rst = New adodb.Recordset
strSQL = Replace(strSQL, "?BATCH?", batch)
rst.Open strSQL, Conn, adOpenKeyset, adLockOptimistic
    ' execute query
    If (Not rst Is Nothing) Then
        If (Not rst.EOF And Not IsNull(rst.Fields(0))) Then
            maxLineNo = rst.Fields(0).Value
        End If
    End If
Navision_GetMaxLineNo = maxLineNo
rst.Close
Set rst = Nothing
Conn.Close
Set Conn = Nothing
End Function

' Go the the External Jrnl Line No Cntrl Table to retrieve the first Line No.
Function Navision_GetBeginLineNo(batch As String) As Long
Dim strSQL As String
strSQL = "SELECT [Beg Line No_] FROM [Hubbard Broadcasting Inc_$External Jrnl Line No Cntrl] WHERE [Journal Batch Name] = '?BATCH?'"
Dim lineNo As Long
Dim rst As adodb.Recordset
Set rst = New adodb.Recordset
Dim Conn As adodb.Connection
Set Conn = New adodb.Connection
Conn.Open ADOconn
lineNo = 0
strSQL = Replace(strSQL, "?BATCH?", batch)
' execute query
rst.Open strSQL, Conn, adOpenKeyset, adLockOptimistic
    If Not rst Is Nothing Then
        If Not rst.EOF And Not IsNull(rst.Fields(0)) Then
            lineNo = rst.Fields(0).Value ' extract line number from first field
        End If
    End If
    Navision_GetBeginLineNo = lineNo
    rst.Close
    Set rst = Nothing
    Conn.Close
    Set Conn = Nothing
End Function
Function GetUserConfirmationToWrite() As Boolean
    Dim msgboxHeader As String
    Dim prompt As String
    Dim buttons As Integer
    Dim msgreturn As Integer
    Dim isConfirmed As Boolean
    
    isConfirmed = False
    
    msgboxHeader = "Writing Journal #" + CStr(ActiveSheet.Range("J3").Value) _
        + " for Division #" + CStr(ActiveSheet.Range("I3").Value)
    
    prompt = "Warning! This option writes all unwritten Journal Entries to the " _
        + "General Ledger system.  Are you sure you want to do this?"
    buttons = vbYesNoCancel + vbExclamation + vbDefaultButton2
    
    msgreturn = MsgBox(prompt, buttons, msgboxHeader)
    
    Select Case msgreturn
    Case vbYes
        isConfirmed = True
    End Select

    GetUserConfirmationToWrite = isConfirmed
End Function




