Attribute VB_Name = "Module1"
Option Explicit

Dim AppObject As New clsApp
    
Sub Init()
'   Called by Workbook_Open
    Set AppObject.AppEvents = Application
End Sub

Sub UpdateLogFile(Wb)
    Dim txt As String
    Dim Fname As String
    On Error Resume Next
    txt = Wb.FullName
    txt = txt & "," & Date & "," & Time
    txt = txt & "," & Application.UserName
    Fname = Application.DefaultFilePath & "\logfile.csv"
    Open Fname For Append As #1
    MsgBox txt
    Print #1, txt
    Close #1
End Sub

Function DefaultFileDirectory()
    DefaultFileDirectory = Application.DefaultFilePath
End Function
