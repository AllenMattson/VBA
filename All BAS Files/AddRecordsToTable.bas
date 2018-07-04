Attribute VB_Name = "AddRecordsToTable"
Sub AddRecordsToTable()
'USING ADO
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
    With rs
        .Open "SELECT * FROM ZillowMaster", sConnString, adOpenKeyset, adLockOptimistic
        'ADD NEW RECORD AND SPECIFY VALUES
        .AddNew
        '''''''''''''''''''''''
        'COLUMN NAMES LIKE THIS---- ![COLUMN_NAME] = THE DATA YOU ARE INSERTING
        ''''''''''''''''''''''
        .MoveFirst 'Moves new record to the first of recordset
        .Close
    End With
    
    Set rs = Nothing
    Set conn = Nothing
End Sub
