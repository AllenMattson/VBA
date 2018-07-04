Attribute VB_Name = "ado_Recordset_to_array"

Dim conn As ADODB.Connection
Dim statement As String
Dim rs As ADODB.Recordset
Dim values As Variant
Dim txt As String
Dim r As Integer
Dim c As Integer

    ' Open the database connection.
    '...

    ' Select the data.
    statement = "SELECT * FROM Books ORDER BY Title, Year"

    ' Get the records.
    Set rs = conn.Execute(statement, , adCmdText)

    ' Load the values into a variant array.
    values = rs.GetRows

    ' Close the recordset and connection.
    rs.Close
    conn.Close

    ' Use the array to build a string
    ' containing the results.
    For r = LBound(values, 2) To UBound(values, 2)
        For c = LBound(values, 1) To UBound(values, 1)
            txt = txt & values(c, r) & ", "
        Next c
        txt = Left$(txt, Len(txt) - 1) & vbCrLf
    Next r

    ' Display the results.
    txtBooks.Text = txt
