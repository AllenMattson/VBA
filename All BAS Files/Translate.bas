Attribute VB_Name = "Translate"
Private pClient As WebClient
Public Property Get Client() As WebClient
    If pClient Is Nothing Then
        Set pClient = New WebClient
        pClient.BaseUrl = "https://www.googleapis.com/language/translate/v2"
    End If
    
    Set Client = pClient
End Property

Public Function Translate(Target As String, Text As String, Optional Source As String = "en") As WebResponse
    Dim Request As New WebRequest
    Request.AddQuerystringParam "key", Credentials.Values("Google")("api_key")
    Request.AddQuerystringParam "target", Target
    Request.AddQuerystringParam "q", Text
    
    Set Translate = Client.Execute(Request)
End Function

Public Sub Test()
    Dim Response As WebResponse
    Set Response = Translate("de", "Hello World")
    
    If Response.StatusCode = WebStatusCode.Ok Then
        Debug.Print "Translation: " & Response.Data("data")("translations")(1)("translatedText")
    Else
        Debug.Print Response.Content
    End If
End Sub


