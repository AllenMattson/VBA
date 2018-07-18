Attribute VB_Name = "Module1"
Option Explicit

Sub RangeToVariant2()
    Dim x As Variant
    Dim r As Long, c As Integer

'   Read the data into the variant
    x = Range("data").Value
    
'   Loop through the variant array
    For r = 1 To UBound(x, 1)
        For c = 1 To UBound(x, 2)
'           Multiply by 2
            x(r, c) = x(r, c) * 2
        Next c
    Next r

'   Transfer the variant back to the sheet
    Range("data") = x
End Sub

