Function WsExist( _
    WsName As String) As Boolean

    Dim Ws As Worksheet

    WsExist = False
       
    For Each Ws In Worksheets
        If Ws.Name = WsName Then
            WsExist = True
            Exit Function
        End If
    Next Ws
    
End Function



Sub SampleUsage()
    
    'Sample usage for WsExist()
    
    Dim WsName as String
    
    WsName = "Test"
    
    If WsExist(WsName) Then
        MsgBox "Worksheet " & WsName & " exists in this workbook."
    Else
        MsgBox "Worksheet " & WsName & " does not exists in this workbook."
    End If
    
End Sub
