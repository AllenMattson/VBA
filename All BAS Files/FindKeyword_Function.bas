Attribute VB_Name = "FindKeyword_Function"
Function FindKeyword( _
    WordList As Range, _
    TextString As Range)

    'Finds keywords in a text string from a user supplied list of words.
    
    Dim Word As Range

    'Test each word in the provided WordList against TextString   '
    For Each Word In WordList
        'If a word is foun, add it to the result
        If InStr(1, TextString, Word, 1) > 0 Then
            If IsEmpty(FindKeyword) Then
                FindKeyword = Word
            Else
                'If more than one word is found, result is deliminated with commas
                FindKeyword = FindKeyword & ", " & Word
            End If
        End If
    Next Word
    
    'Reuturn "" when no words arefound in the TextString
    If FindKeyword = 0 Then
        FindKeyword = ""
    End If
    
End Function


