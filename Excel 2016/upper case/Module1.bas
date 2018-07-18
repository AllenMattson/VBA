Attribute VB_Name = "Module1"
Function UPCASE(instring As String) As String
'   Converts its argument to all uppercase.
    Dim StringLength As Long
    Dim i As Long
    Dim ASCIIVal As Long
    Dim CharVal As Long
    
    StringLength = Len(instring)
    UPCASE = instring
    For i = 1 To StringLength
        ASCIIVal = Asc(Mid(instring, i, 1))
        CharVal = 0
        If ASCIIVal >= 97 And ASCIIVal <= 122 Then
            CharVal = -32
            Mid(UPCASE, i, 1) = Chr(ASCIIVal + CharVal)
        End If
    Next i
End Function

