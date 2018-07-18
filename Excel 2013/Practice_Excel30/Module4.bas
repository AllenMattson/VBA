Attribute VB_Name = "Module4"
Option Explicit

Public Declare Function GetVersionEx Lib "kernel32" Alias _
    "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) _
    As Long
  
Public Declare Function GetWindowsDirectory _
    Lib "kernel32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) _
    As Long
    
Public Declare Function GetUserName _
    Lib "advapi32.dll" Alias "GetUserNameA" _
    (ByVal lpBuffer As String, nSize As Long) As Long

Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Sub OpSysInfo()
    Dim os As OSVERSIONINFO
    Dim osVer As String
    
    os.dwOSVersionInfoSize = Len(os)
    GetVersionEx os
    osVer = os.dwMajorVersion & "." & os.dwMinorVersion
    Debug.Print "Windows Version = " & osVer
    Debug.Print "Windows Build Number = " & os.dwBuildNumber
    Debug.Print "Windows Platform ID = " & os.dwPlatformId
    Debug.Print "Additional info = " & os.szCSDVersion
End Sub

Sub PathToWinDir()
    Dim strWinDir As String
    Dim lngLen As Long
       
    strWinDir = String(255, 0)
    lngLen = GetWindowsDirectory(strWinDir, Len(strWinDir))
    strWinDir = Left(strWinDir, lngLen)
    MsgBox "Windows folder: " & strWinDir
End Sub

Function LoggedOnUserName() As String
    Dim strBuffer As String * 255
    Dim strLen As Long
 
    strLen = Len(strBuffer)
    GetUserName strBuffer, strLen
 
    If strLen > 0 Then
        LoggedOnUserName = Left$(strBuffer, strLen - 1)
    End If
    
    MsgBox LoggedOnUserName
End Function




