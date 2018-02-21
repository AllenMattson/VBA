Attribute VB_Name = "StaticArrays"
Option Explicit

' start indexing array elements at 1
Option Base 1

Sub FavoriteCities()
  'now declare the array
  Dim cities(6) As String

  'assign the values to array elements
  cities(1) = "Baltimore"
  cities(2) = "Atlanta"
  cities(3) = "Boston"
  cities(4) = "San Diego"
  cities(5) = "New York"
  cities(6) = "Denver"

  'display the list of cities
  MsgBox cities(1) & Chr(13) & cities(2) & Chr(13) _
      & cities(3) & Chr(13) & cities(4) & Chr(13) _
      & cities(5) & Chr(13) & cities(6)
End Sub

Sub FavoriteCities2()
  'now declare the array
  Dim cities(6) As String
  Dim city As Variant

  'assign the values to array elements
  cities(1) = "Baltimore"
  cities(2) = "Atlanta"
  cities(3) = "Boston"
  cities(4) = "San Diego"
  cities(5) = "New York"
  cities(6) = "Denver"

  'display the list of cities in separate messages
  For Each city In cities
      MsgBox city
  Next
End Sub

Sub FavoriteCities3()
  'now declare the array
  Dim cities(6) As String

  'assign the values to array elements
  cities(1) = "Baltimore"
  cities(2) = "Atlanta"
  cities(3) = "Boston"
  cities(4) = "San Diego"
  cities(5) = "New York"
  cities(6) = "Denver"
    
  'call another procedure and pass the array as argument
  Hello cities()
End Sub

Sub Hello(cities() As String)
  Dim counter As Integer

  For counter = LBound(cities()) To UBound(cities())
      MsgBox "Hello, " & cities(counter) & "!"
  Next
End Sub

Sub Lotto()
  Const spins = 6
  Const minNum = 1
  Const maxNum = 54

  Dim t As Integer          ' looping variable in outer loop
  Dim i As Integer          ' looping variable in inner loop
  Dim myNumbers As String       ' string to hold all picks
  Dim lucky(spins) As String    ' array to hold generated picks

  myNumbers = ""

  For t = 1 To spins
    Randomize
    lucky(t) = Int((maxNum - minNum + 1) * Rnd) + minNum

    ' see if this number was drawn before
    For i = 1 To (t - 1)
      If lucky(t) = lucky(i) Then
          lucky(t) = Int((maxNum - minNum + 1) * Rnd) + minNum
          i = 0
      End If
    Next i
    MsgBox "Lucky number is " & lucky(t)
    myNumbers = myNumbers & " - " & lucky(t)
  Next t

  MsgBox "Lucky numbers are " & myNumbers
End Sub

Sub Exchange()
  Dim t As String
  Dim r As String
  Dim Ex(3, 3) As Variant

    t = Chr(9)  ' tab
    r = Chr(13) ' Enter

    Ex(1, 1) = "Japan"
    Ex(1, 2) = "Yen"
    Ex(1, 3) = 104.57
    Ex(2, 1) = "Mexico"
    Ex(2, 2) = "Peso"
    Ex(2, 3) = 11.2085
    Ex(3, 1) = "Canada"
    Ex(3, 2) = "Dollar"
    Ex(3, 3) = 1.2028
    MsgBox "Country " & t & t & "Currency" & t & "per US$" _
      & r & r _
      & Ex(1, 1) & t & t & Ex(1, 2) & t & Ex(1, 3) & r _
      & Ex(2, 1) & t & t & Ex(2, 2) & t & Ex(2, 3) & r _
      & Ex(3, 1) & t & t & Ex(3, 2) & t & Ex(3, 3) & r & r _
      & "* Sample Exchange Rates for Demonstration Only", , _
      "Exchange"
End Sub


