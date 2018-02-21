Attribute VB_Name = "Module1"
Option Explicit

Sub GetName()
    Dim UserName As String
    Dim FirstSpace As Integer
    Do Until UserName <> ""
        UserName = InputBox("Enter your full name: ", _
            "Identify Yourself")
    Loop
    FirstSpace = InStr(UserName, " ")
    If FirstSpace <> 0 Then
        UserName = Left(UserName, FirstSpace - 1)
    End If
    MsgBox "Hello " & UserName
End Sub


Sub GetWord()
    Dim TheWord As String
    Dim p As String
    Dim t As String
    p = Range("A1")
    t = "What's the missing word?"
    TheWord = InputBox(prompt:=p, Title:=t)
    If UCase(TheWord) = "BATTLEFIELD" Then
        MsgBox "Correct."
    Else
        MsgBox "That is incorrect."
    End If
End Sub
