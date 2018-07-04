Attribute VB_Name = "FTP_Get_Put"
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName _
As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" _
Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, _
ByVal nServerPort As Integer, ByVal sUsername As String, _
ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function FTPGETFile Lib "wininet.dll" Alias "FtpGetFileA" _
(ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, _
ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, _
ByVal dwContext As Long) As Boolean
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
ByVal lpszRemoteFile As String, ByVal dwFlags As Long, _
ByVal dwContext As Long) As Boolean
Private Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer
'set up a subroutine to transfer a file from your chosen FTP server to your local PC:
Sub GetFile()
Dim MyConn, MyINet, Chk As Boolean
Chk = False
MyINet = InternetOpen("MyFTP", 1, vbNullString, vbNullString, 0)
If MyINet > 0 Then
MyConn = InternetConnect(MyINet, "MyOrg.MyServer.net", 21, "MyUserID", "MyPassword", 1, 0, 0)
If MyConn > 0 Then
Chk = FTPGETFile(MyConn, "MyFolder/MyFileName.txt", _
CurrentProject.Path & "\MyFileName.txt", 0, 0, 1, 0)
InternetCloseHandle MyConn
End If
InternetCloseHandle MyINet
End If
If (Chk) Then
MsgBox "File downloaded"
Else
MsgBox "FTP Error"
End If
End Sub
'files can be transferred from the local PC to the FTP server, provided that your user ID has the necessary permissions
Sub PutFile()
Dim MyConn, MyINet, Chk As Boolean
Chk = False
MyINet = InternetOpen("MyFTP", 1, vbNullString, vbNullString, 0)
If MyINet > 0 Then
MyConn = InternetConnect(MyINet, "MyOrg.MyServer.net", 21, "MyUserID", "MyPassword", 1, 0, 0)
If MyConn > 0 Then
Chk = FtpPutFile(MyConn, CurrentProject.Path & "\test.txt", " MyFolder/MyFileName.txt", 1, 0)
InternetCloseHandle MyConn
End If
InternetCloseHandle MyINet
End If
If (Chk) Then
MsgBox "File uploaded"
Else
MsgBox "FTP Error"
End If
End Sub
