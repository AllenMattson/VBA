Attribute VB_Name = "Errors_in_Arrays"
Option Explicit

Sub Zoo1()
  ' this procedure triggers an error "Subscript out of range"
  Dim zoo(1 To 3) As String
  Dim i As Integer
  Dim response As String
  i = 1

  Do
    response = InputBox("Enter a name of animal:")
    zoo(i) = response
    i = i + 1
  Loop Until response = ""
End Sub


Sub Zoo2()
      ' this procedure avoids the error "Subscript out of range"
      Dim zoo(1 To 3) As String
      Dim i As Integer
      Dim response As String
      i = 1

      Do While i >= LBound(zoo) And i <= UBound(zoo)
        response = InputBox("Enter a name of animal:")
        If response = "" Then Exit Sub
        zoo(i) = response
        i = i + 1
      Loop
      
      i = 0
      For i = LBound(zoo) To UBound(zoo)
        MsgBox zoo(i)
      Next
End Sub


