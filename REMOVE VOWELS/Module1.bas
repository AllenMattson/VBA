Attribute VB_Name = "Module1"
Function RemoveVowels(txt) As String
Attribute RemoveVowels.VB_Description = "Returns the argument, with all vowels removed."
Attribute RemoveVowels.VB_HelpID = 50100
Attribute RemoveVowels.VB_ProcData.VB_Invoke_Func = " \n14"
' Removes all vowels from the Txt argument
    Dim i As Long
    RemoveVowels = ""
    For i = 1 To Len(txt)
        If Not ucase(Mid(txt, i, 1)) Like "[AEIOU]" Then
            RemoveVowels = RemoveVowels & Mid(txt, i, 1)
        End If
    Next i
End Function


Function RemoveVowels2(txt) As String
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
    RemoveVowels2 = TempString
End Function


Sub ZapTheVowels()
     Dim UserInput As String
     UserInput = InputBox("Enter some text:")
     MsgBox RemoveVowels(UserInput), vbInformation, UserInput
End Sub

