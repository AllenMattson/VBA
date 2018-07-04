' Written: April 29, 2012
' Author:  Leith Ross
' Summary: Returns the status for a URL along with the Page Source HTML text.

Public PageSource As String
Public httpRequest As Object

Function GetURLStatus(ByVal URL As String, Optional AllowRedirects As Boolean)

    Const WinHttpRequestOption_EnableRedirects = 6
  
  
        If httpRequest Is Nothing Then
            On Error Resume Next
                Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
                If httpRequest Is Nothing Then
                    Set httpRequest = CreateObject("WinHttp.WinHttpRequest.5")
                End If
            Err.Clear
            On Error GoTo 0
        End If

        ' Control if the URL being queried is allowed to redirect.
          httpRequest.Option(WinHttpRequestOption_EnableRedirects) = AllowRedirects

        ' Clear any pervious web page source information
          PageSource = ""
    
        ' Add protocol if missing
          If InStr(1, URL, "://") = 0 Then
             URL = "http://" & URL
          End If

             ' Launch the HTTP httpRequest synchronously
               On Error Resume Next
                  httpRequest.Open "GET", URL, False
                  If Err.Number <> 0 Then
                   ' Handle connection errors
                     GetURLStatus = Err.Description
                     Err.Clear
                     Exit Function
                  End If
               On Error GoTo 0
           
             ' Send the http httpRequest for server status
               On Error Resume Next
                  httpRequest.Send
                  httpRequest.WaitForResponse
                  If Err.Number <> 0 Then
                   ' Handle server errors
                     PageSource = "Error"
                     GetURLStatus = Err.Description
                     Err.Clear
                  Else
                   ' Show HTTP response info
                     GetURLStatus = httpRequest.Status & " - " & httpRequest.StatusText
                   ' Save the web page text
                     PageSource = httpRequest.ResponseText
                  End If
               On Error GoTo 0
            
End Function

Sub ValidateURLs()

    Dim Cell As Range
    Dim Rng As Range
    Dim RngEnd As Range
    Dim Status As String
    Dim Wks As Worksheet
    
        Set Wks = ActiveSheet
        Set Rng = Wks.Range("F2")
        
        Set RngEnd = Wks.Cells(Rows.Count, Rng.Column).End(xlUp)
        If RngEnd.Row < Rng.Row Then Exit Sub Else Set Rng = Wks.Range(Rng, RngEnd)
        
            For Each Cell In Rng
                Status = GetURLStatus(Cell)
                If Status <> "200 - OK" Then
                   Cell = Status
                End If
            Next Cell
        
End Sub