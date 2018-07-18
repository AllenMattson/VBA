Attribute VB_Name = "Module1"
Option Explicit

Sub GetName()
    Dim UserName As String
    Dim FirstSpace As Long
    Do Until Len(UserName) > 0
        UserName = InputBox("Enter your full name: ", _
            "Identify Yourself")
    Loop
    FirstSpace = InStr(UserName, Space(1))
    If FirstSpace > 0 Then
        UserName = Left$(UserName, FirstSpace - 1)
    End If
    MsgBox "Hello " & UserName
End Sub


Sub GetWord()
    Dim TheWord As String
    Dim Prompt As String
    Dim Title As String
    Prompt = Range("A1")
    Title = "What's the missing word?"
    TheWord = InputBox(Prompt:=Prompt, Title:=Title)
    If UCase(TheWord) = "BATTLEFIELD" Then
        MsgBox "Correct."
    Else
        MsgBox "That is incorrect."
    End If
End Sub

Sub GetValue()
    Dim Monthly As String
    Monthly = InputBox("Enter your monthly salary:")
    If Len(Monthly) > 0 And IsNumeric(Monthly) Then
        MsgBox "Annualized: " & Monthly * 12
    Else
        MsgBox "Invalid input"
    End If
End Sub
