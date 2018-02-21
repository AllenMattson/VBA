Attribute VB_Name = "Module1"
Function UpCase(instring As String) As String
'   Converts its argument to all uppercase.
    Dim StringLength As Integer
    Dim i As Integer
    Dim ASCIIVal As Integer
    Dim CharVal As Integer
    
    StringLength = Len(instring)
    UpCase = instring
    For i = 1 To StringLength
        ASCIIVal = Asc(Mid(instring, i, 1))
        CharVal = 0
        If ASCIIVal >= 97 And ASCIIVal <= 122 Then
            CharVal = -32
            Mid(UpCase, i, 1) = Chr(ASCIIVal + CharVal)
        End If
    Next i
End Function

Function UpCase2(instring As String) As String
    UpCase2 = UCase(instring)
End Function
