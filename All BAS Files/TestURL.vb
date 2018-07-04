Private Function TestURL(strUrl As String) As Boolean
    
    'Requires Microsoft WinHTTP Services, Version 5.1
    
    'Function returns TRUE is URL is found; FALSE if not found
    'Default value of TestURL is FALSE
    
    Dim oURL As New WinHttpRequest

    On Error GoTo TestURL_Err
    
    With oURL
        .Open "GET", strUrl, False
        .send
        If .Status = 200 Then          '200 indicates resource was retrieved
            TestURL = True
        End If
    End With
     
TestURL_Err:
    'Invalid resources cause an error; Function returns default value
    
    Set oURL = Nothing                 'Clean up object

End Function
