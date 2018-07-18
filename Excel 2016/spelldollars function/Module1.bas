Attribute VB_Name = "Module1"
Function SPELLDOLLARS(cell) As Variant

    Dim Dollars As String
    Dim Cents As String
    Dim TextLen As Integer
    Dim Temp As String
    Dim Pos As Integer
    Dim iHundreds As Integer
    Dim iTens As Integer
    Dim iOnes As Integer
    Dim Units(2 To 5) As String
    Dim bHit As Boolean
    Dim Ones As Variant
    Dim Teens As Variant
    Dim Tens As Variant
    Dim NegFlag As Boolean

'   Is it a non-number?
    If Not IsNumeric(cell) Then
        SPELLDOLLARS = CVErr(xlErrValue)
        Exit Function
    End If

'   Is it negative?
    If cell < 0 Then
        NegFlag = True
        cell = Abs(cell)
    End If
    
    Dollars = Format(cell, "###0.00")
    TextLen = Len(Dollars) - 3

'   Is it too large?
    If TextLen > 15 Then
        SPELLDOLLARS = CVErr(xlErrNum)
        Exit Function
    End If

'   Do the cents part
    Cents = Right(Dollars, 2) & "/100 Dollars"
    If cell < 1 Then
        SPELLDOLLARS = Cents
        Exit Function
    End If

    Dollars = Left(Dollars, TextLen)

    Ones = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine")
    Teens = Array("Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
    Tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")

    Units(2) = "Thousand"
    Units(3) = "Million"
    Units(4) = "Billion"
    Units(5) = "Trillion"

    Temp = ""

    For Pos = 15 To 3 Step -3
        If TextLen >= Pos - 2 Then
            bHit = False
            If TextLen >= Pos Then
                iHundreds = Asc(Mid$(Dollars, TextLen - Pos + 1, 1)) - 48
                If iHundreds > 0 Then
                    Temp = Temp & " " & Ones(iHundreds) & " Hundred"
                    bHit = True
                End If
        End If
        iTens = 0
        iOnes = 0

        If TextLen >= Pos - 1 Then
            iTens = Asc(Mid$(Dollars, TextLen - Pos + 2, 1)) - 48
        End If

        If TextLen >= Pos - 2 Then
            iOnes = Asc(Mid$(Dollars, TextLen - Pos + 3, 1)) - 48
        End If

        If iTens = 1 Then
            Temp = Temp & " " & Teens(iOnes)
            bHit = True
        Else
            If iTens >= 2 Then
                Temp = Temp & " " & Tens(iTens)
                bHit = True
            End If
            If iOnes > 0 Then
                If iTens >= 2 Then
                    Temp = Temp & "-"
                Else
                    Temp = Temp & " "
                End If
                Temp = Temp & Ones(iOnes)
                bHit = True
            End If
        End If
        If bHit And Pos > 3 Then
            Temp = Temp & " " & Units(Pos \ 3)
        End If
    End If
    Next Pos

  SPELLDOLLARS = Trim(Temp) & " and " & Cents
  If NegFlag Then SPELLDOLLARS = "(" & SPELLDOLLARS & ")"
End Function

