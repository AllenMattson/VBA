Attribute VB_Name = "Module1"
Function REMOVEVOWELS(txt) As String
Attribute REMOVEVOWELS.VB_Description = "Returns the argument, with all vowels removed."
Attribute REMOVEVOWELS.VB_HelpID = 50100
Attribute REMOVEVOWELS.VB_ProcData.VB_Invoke_Func = " \n14"
' Removes all vowels from the Txt argument
    Dim i As Long
    REMOVEVOWELS = ""
    For i = 1 To Len(txt)
        If Not ucase(Mid(txt, i, 1)) Like "[AEIOU]" Then
            REMOVEVOWELS = REMOVEVOWELS & Mid(txt, i, 1)
        End If
    Next i
End Function


Function REMOVEVOWELS2(txt) As String
' Removes all vowels from the Txt argument
    Dim i As Long
    Dim TempString As String
    TempString = ""
    For i = 1 To Len(txt)
        Select Case ucase(Mid(txt, i, 1))
            Case "A", "E", "I", "O", "U"
                'Do nothing
            Case Else
                TempString = TempString & Mid(txt, i, 1)
        End Select
    Next i
    REMOVEVOWELS2 = TempString
End Function


Sub ZapTheVowels()
     Dim UserInput As String
     UserInput = InputBox("Enter some text:")
     MsgBox REMOVEVOWELS(UserInput), vbInformation, UserInput
End Sub

