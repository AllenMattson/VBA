Attribute VB_Name = "SeeIfInternet_IsConnected"
Public Declare Function InternetGetConnectedState _
Lib "wininet.dll" (lpdwFlags As Long, _
ByVal dwReserved As Long) As Boolean
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" Alias "InternetGetConnectedStateExA" ( _
ByRef lpdwFlags As Long, _
ByVal lpszConnectionName As String, _
ByVal dwNameLen As Long, _
ByVal dwReserved As Long) As Long 'Local system uses a modem to connect to the Internet.
Private Const INTERNET_CONNECTION_MODEM As Long = &H1 'Local system uses a LAN to connect to the Internet.
Private Const INTERNET_CONNECTION_LAN As Long = &H2 'Local system uses a proxy server to connect to the Internet.
Private Const INTERNET_CONNECTION_PROXY As Long = &H4
'Read more at http://vbadud.blogspot.com/2010/05/how-to-check-internet-connectivity.html#T5bViOHs8uzoro98.99
Function IsConnected() As Boolean

Dim Stat As Long
IsConnected = (InternetGetConnectedState(Stat, 0&) <> 0)
    If IsConnected And INTERNET_CONNECTION_LAN Then
        MsgBox "Lan Connection"
    ElseIf IsConnected And INTERNET_CONNECTION_MODEM Then
        MsgBox "Modem Connection"
    ElseIf IsConnected And INTERNET_CONNECTION_PROXY Then
        MsgBox "Proxy"
    End If
End Function
'Read more at http://vbadud.blogspot.com/2010/05/how-to-check-internet-connectivity.html#T5bViOHs8uzoro98.99
