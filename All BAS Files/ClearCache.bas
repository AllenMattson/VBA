Attribute VB_Name = "ClearCache"
'In Tools > References, add reference to "Microsoft XML, vX.X" before running.
Sub subClearCache()
    
    ' force browser to clear cache
    myURL = "http://172.16.50.250/blackberry/BBTESTB01.pgm"
    Dim oHttp As New MSXML2.XMLHTTP
    oHttp.Open "POST", myURL, False
    oHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    oHttp.setRequestHeader "Cache-Control", "no-cache"
    oHttp.setRequestHeader "PragmaoHttp", "no-cache"
    oHttp.setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
    oHttp.setRequestHeader "Authorization", "Basic " & Base64EncodedUsernamePassword
    oHttp.send "PostArg1=PostArg1Value"
    Result = oHttp.responseText
    
End Sub
