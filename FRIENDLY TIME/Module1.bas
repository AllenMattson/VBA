Attribute VB_Name = "Module1"
Function FT(t1, t2)
    Dim SDif As Double, DDif As Double
    
    If Not (IsDate(t1) And IsDate(t2)) Then
        FT = CVErr(xlErrValue)
        Exit Function
    End If
    
    DDif = Abs(t2 - t1)
    SDif = DDif * 24 * 60 * 60
    
    If DDif < 1 Then
       If SDif < 10 Then FT = "Just now": Exit Function
       If SDif < 60 Then FT = SDif & " seconds ago": Exit Function
       If SDif < 120 Then FT = "a minute ago": Exit Function
       If SDif < 3600 Then FT = Round(SDif / 60, 0) & "minutes ago": Exit Function
       If SDif < 7200 Then FT = "An hour ago": Exit Function
       If SDif < 86400 Then FT = Round(SDif / 3600, 0) & " hours ago": Exit Function
    End If
    If DDif = 1 Then FT = "Yesterday": Exit Function
    If DDif < 7 Then FT = Round(DDif, 0) & " days ago": Exit Function
    If DDif < 31 Then FT = Round(DDif / 7, 0) & " weeks ago": Exit Function
    If DDif < 365 Then FT = Round(DDif / 30, 0) & " months ago": Exit Function
    FT = Round(DDif / 365, 0) & " years ago"
End Function
