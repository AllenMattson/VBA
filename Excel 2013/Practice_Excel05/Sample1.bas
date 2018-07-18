Attribute VB_Name = "Sample1"
Option Explicit


Public Function SumItUp(m, n)
  SumItUp = m + n
End Function


Sub RunSumItUp()
  Dim m As Single, n As Single
  m = 370000
  n = 3459.77

  Debug.Print SumItUp(m, n)
  MsgBox "Open the Immediate Window to see the result."
End Sub


Sub NumOfCharacters()
  Dim f As Integer
  Dim l As Integer

  f = Len(InputBox("Enter first name:"))
  l = Len(InputBox("Enter last name:"))
  MsgBox SumItUp(f, l)
End Sub

