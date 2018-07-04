Attribute VB_Name = "Facebook"
Dim pClient As WebClient
Public Property Get Client() As WebClient
    If pClient Is Nothing Then
        Set pClient = New WebClient
        pClient.BaseUrl = "graph.facebook.com"
        
        Dim Auth As New FacebookAuthenticator
        Auth.Setup CStr(Credentials.Values("Facebook")("id")), CStr(Credentials.Values("Facebook")("secret"))
        Auth.AddScope "user_location"
        Auth.Login
        
        Set pClient.Authenticator = Auth
        
        ' For testing only
        Dim TempToken As String
        TempToken = Auth.GetToken(pClient)
        Debug.Print "Success! " & TempToken
    End If
    
    Set Client = pClient
    
    ' TEMP No caching
    Set pClient = Nothing
End Property

Public Sub Test()
    Dim TestClient As WebClient
    Set TestClient = Client
End Sub
