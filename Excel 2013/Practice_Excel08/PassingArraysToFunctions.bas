Attribute VB_Name = "PassingArraysToFunctions"
Option Explicit

Sub ManipulateArray()
      Dim countries(1 To 6) As Variant
      Dim countriesUCase As Variant
      Dim i As Integer
      Dim r As Integer

      ' assign the values to array elements
      countries(1) = "Bulgaria"
      countries(2) = "Argentina"
      countries(3) = "Brazil"
      countries(4) = "Sweden"
      countries(5) = "New Zealand"
      countries(6) = "Denmark"
    
      countriesUCase = ArrayToUCase(countries)
    
      r = 1 'row counter
    
      With ActiveSheet
        For i = 1 To 6
            Cells(r, 1).Value = countriesUCase(i)
            Cells(r, 2).Value = countries(i)
            r = r + 1
        Next i
      End With
    End Sub

    Public Function ArrayToUCase(ByVal myValues As Variant) _
        As String()
      Dim i As Integer
      Dim Temp() As String
      If IsArray(myValues) Then
        ReDim Temp(LBound(myValues) To UBound(myValues))
        For i = LBound(myValues) To UBound(myValues)
            Temp(i) = CStr(UCase(myValues(i)))
        Next i
        ArrayToUCase = Temp
      End If
    End Function

