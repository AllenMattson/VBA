Attribute VB_Name = "ParameterArrays"
Option Explicit

Function AddMultipleArgs(ParamArray myNumbers() As Variant)
  Dim mySum As Single
  Dim myValue As Variant
  For Each myValue In myNumbers
    mySum = mySum + myValue
  Next
  AddMultipleArgs = mySum
End Function

