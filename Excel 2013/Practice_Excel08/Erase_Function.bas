Attribute VB_Name = "Erase_Function"
Option Explicit

' start indexing array elements at 1
Option Base 1

Sub FunCities()
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
Erase cities

' show all that was were erased
MsgBox cities(1) & Chr(13) & cities(2) & Chr(13) _
  & cities(3) & Chr(13) & cities(4) & Chr(13) _
  & cities(5)
End Sub


