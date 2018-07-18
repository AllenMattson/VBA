Attribute VB_Name = "Module1"
Option Explicit
   
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function FindExecutableA Lib "shell32.dll" _
        (ByVal lpFile As String, ByVal lpDirectory As String, _
        ByVal lpResult As String) As Long
#Else
    Private Declare Function FindExecutableA Lib "shell32.dll" _
        (ByVal lpFile As String, ByVal lpDirectory As String, _
        ByVal lpResult As String) As Long
#End If


Function GetExecutable(strFile As String) As String
    Dim strPath As String
    Dim intLen As Integer
    strPath = Space(255)
    intLen = FindExecutableA(strFile, "\", strPath)
    GetExecutable = Trim(strPath)
End Function

Sub GetFileName()
    Dim fname As String
    fname = Application.GetOpenFilename
    MsgBox "The executable file is " & vbCrLf & vbCrLf & GetExecutable(fname), vbInformation, fname
End Sub

