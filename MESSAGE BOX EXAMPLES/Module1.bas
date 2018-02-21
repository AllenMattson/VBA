Attribute VB_Name = "Module1"
Option Explicit

Sub LongText()
    MsgBox Range("A1"), , "Gettysburg Address"
End Sub

Sub ShowRange()
    Dim Msg As String
    Dim r As Long, c As Long
    Msg = ""
    For r = 1 To 13
        For c = 1 To 3
            Msg = Msg & Cells(r, c).Text
            If c <> 3 Then Msg = Msg & vbTab
        Next c
        Msg = Msg & vbNewLine
    Next r
    MsgBox Msg
End Sub

Sub RunProc()
    
    If ContinueProcedure Then
        'do stuff
    End If
    
End Sub

Private Function ContinueProcedure() As Boolean
    Dim Config As Long
    Dim Ans As Long
    Config = vbYesNo + vbQuestion + vbDefaultButton2
    Ans = MsgBox("An error occurred. Continue?", Config)
    If Ans = vbYes Then ContinueProcedure = True _
        Else ContinueProcedure = False
End Function


Sub test()
    
    Dim x
    
    x = VBA.MsgBox("x",
End Sub
