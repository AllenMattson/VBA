Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 And Win64 Then
  Declare PtrSafe Function GetWindowsDirectoryA Lib "kernel32" _
  (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#Else
  Declare Function GetWindowsDirectoryA Lib "kernel32" _
  (ByVal lpBuffer As String, ByVal nSize As Long) As Long
#End If



Sub ShowWindowsDir()
    Dim WinPath As String * 255
    Dim WinDir As String
    WinPath = Space(255)
    WinDir = Left(WinPath, GetWindowsDirectoryA(WinPath, Len(WinPath)))
    MsgBox WinDir, vbInformation, "Windows Directory"
End Sub


Function WINDOWSDIR() As String
'   Returns the Windows directory
    Dim WinPath As String * 255
    WinPath = Space(255)
    WINDOWSDIR = Left(WinPath, GetWindowsDirectoryA(WinPath, Len(WinPath)))
End Function

