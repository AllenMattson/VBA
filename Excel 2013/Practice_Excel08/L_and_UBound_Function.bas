Attribute VB_Name = "L_and_UBound_Function"
Option Explicit

Sub FunCities2()
  ' declare the array
  Dim cities(1 To 5) As String

  ' assign the values to array elements
  cities(1) = "Las Vegas"
  cities(2) = "Orlando"
  cities(3) = "Atlantic City"
  cities(4) = "New York"
  cities(5) = "San Francisco"

    ' display the list of cities
  MsgBox cities(1) & Chr(13) & cities(2) & Chr(13) _
    & cities(3) & Chr(13) & cities(4) & Chr(13) _
    & cities(5)
  ' display the array bounds
  MsgBox "The lower bound: " & LBound(cities) & Chr(13) _
    & "The upper bound: " & UBound(cities)
End Sub


