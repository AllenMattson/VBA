Attribute VB_Name = "ModFunctions"
Option Explicit

Function DECIMAL2RGB(ColorVal) As Variant
'   Converts a color value to an RGB triplet
'   Returns a 3-element variant array
    DECIMAL2RGB = Array(ColorVal \ 256 ^ 0 And 255, ColorVal \ 256 ^ 1 And 255, ColorVal \ 256 ^ 2 And 255)
End Function

Function RGB2DECIMAL(R, G, B)
    RGB2DECIMAL = RGB(R, G, B)
End Function

Function HSL2RGB(H, S, L) As Variant
'   Converts an HSL triplet (0-255) to an RGB triplet
'   Returns a 3-element variant array
    
    Dim temp1 As Single, temp2 As Single
    Dim Rtemp3 As Single, Gtemp3 As Single, Btemp3 As Single
    Dim R As Single, G As Single, B As Single
    
    If S = 0 Then
        HSL2RGB = Array(L, L, L)
        Exit Function
    End If
    
    H = H / 255
    S = S / 255
    L = L / 255
    
    If L < 0.5 Then
        temp2 = L * (1 + S)
    Else
        temp2 = L + S - L * S
    End If
    
    temp1 = 2 * L - temp2
    
    Rtemp3 = H + 1 / 3
    If Rtemp3 < 0 Then Rtemp3 = Rtemp3 + 1
    If Rtemp3 > 1 Then Rtemp3 = Rtemp3 - 1
    Gtemp3 = H
    If Gtemp3 < 0 Then Gtemp3 = Gtemp3 + 1
    If Gtemp3 > 1 Then Gtemp3 = Gtemp3 - 1
    Btemp3 = H - 1 / 3
    If Btemp3 < 0 Then Btemp3 = Btemp3 + 1
    If Btemp3 > 1 Then Btemp3 = Btemp3 - 1
    
    'Red
    
    If 6 * Rtemp3 < 1 Then
       R = temp1 + (temp2 - temp1) * 6 * Rtemp3
    Else
       If 2 * Rtemp3 < 1 Then
          R = temp2
       Else
          If 3 * Rtemp3 < 2 Then
             R = temp1 + (temp2 - temp1) * ((2 / 3) - Rtemp3) * 6
          Else
             R = temp1
           End If
        End If
     End If
    

    'Green
    If 6 * Gtemp3 < 1 Then
        G = temp1 + (temp2 - temp1) * 6 * Gtemp3
    Else
        If 2 * Gtemp3 < 1 Then
            G = temp2
        Else
            If 3 * Gtemp3 < 2 Then
                G = temp1 + (temp2 - temp1) * ((2 / 3) - Gtemp3) * 6
            Else
                G = temp1
            End If
        End If
    End If
    
    'Blue
    If 6 * Btemp3 < 1 Then
        B = temp1 + (temp2 - temp1) * 6 * Btemp3
    Else
        If 2 * Btemp3 < 1 Then
            B = temp2
        Else
            If 3 * Btemp3 < 2 Then
                B = temp1 + (temp2 - temp1) * ((2 / 3) - Btemp3) * 6
            Else
                B = temp1
            End If
        End If
    End If
    
    HSL2RGB = Array(Int(R * 255), Int(G * 255), Int(B * 255))

End Function

Function RGB2HSL(R, G, B) As Variant
'   Converts an RGB triplet to an HSL triplet(0-255)
'   Returns a 3-element variant array
    Dim sMax As Single, sMin As Single
    Dim H As Single, S As Single, L As Single

    R = R / 255
    G = G / 255
    B = B / 255
    sMax = Application.Max(R, Application.Max(G, B))
    sMin = Application.Min(R, Application.Min(G, B))
    L = (sMax + sMin) / 2
    If sMax = sMin Then
        S = 0
        H = 0
        L = L * 255
        RGB2HSL = Array(Round(H, 0), Round(S, 0), Round(L, 0))
        Exit Function
    End If
    If L < 0.5 Then
        S = (sMax - sMin) / (sMax + sMin)
    Else
        S = (sMax - sMin) / (2 - sMax - sMin)
    End If
    If S < 0 Then S = 0
    If R = sMax Then H = (G - B) / (sMax - sMin)
    If G = sMax Then H = 2 + (B - R) / (sMax - sMin)
    If B = sMax Then H = 4 + (R - G) / (sMax - sMin)
    H = H * 42.5
    If H < 0 Then H = H + 255
    S = S * 255
    L = L * 255
    RGB2HSL = Array(Round(H, 0), Round(S, 0), Round(L, 0))
End Function

Function DECIMAL2HSL(ColorVal) As Variant
'   Converts a color value to HSL (0-255)
'   Returns a 3-element variant array
    Dim sMax As Single, sMin As Single
    Dim R As Single, G As Single, B As Single
    Dim H As Single, S As Single, L As Single

    R = ColorVal \ 256 ^ 0 And 255
    G = ColorVal \ 256 ^ 1 And 255
    B = ColorVal \ 256 ^ 2 And 255

    R = R / 255
    G = G / 255
    B = B / 255
    sMax = Application.Max(R, Application.Max(G, B))
    sMin = Application.Min(R, Application.Min(G, B))
    L = (sMax + sMin) / 2
    If sMax = sMin Then
        S = 0
        H = 0
        L = L * 255
        DECIMAL2HSL = Array(Round(H, 0), Round(S, 0), Round(L, 0))
        Exit Function
    End If
    If L < 0.5 Then
        S = (sMax - sMin) / (sMax + sMin)
    Else
        S = (sMax - sMin) / (2 - sMax - sMin)
    End If
    If S < 0 Then S = 0
    If R = sMax Then H = (G - B) / (sMax - sMin)
    If G = sMax Then H = 2 + (B - R) / (sMax - sMin)
    If B = sMax Then H = 4 + (R - G) / (sMax - sMin)
    H = H * 42.5
    If H < 0 Then H = H + 255
    S = S * 255
    L = L * 255
    DECIMAL2HSL = Array(Round(H, 0), Round(S, 0), Round(L, 0))
End Function


Function Round(alpha, beta) As Long
   Round = WorksheetFunction.Round(alpha, beta)
End Function


