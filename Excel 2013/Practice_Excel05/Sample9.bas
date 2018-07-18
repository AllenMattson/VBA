Attribute VB_Name = "Sample9"
Option Explicit

Sub AboutUserMaster()
  Dim first As String, last As String, full As String

  Call GetUserName(full)

  first = GetFirst(full)
  last = GetLast(full)
  Call DisplayLastFirst(first, last)
End Sub

Sub GetUserName(fullName As String)
  fullName = InputBox("Enter first and last name:")
End Sub

Function GetFirst(fullName As String)
  Dim space As Integer

  space = InStr(fullName, " ")

  GetFirst = Left(fullName, space - 1)
End Function

  Dim space As Integer

  space = InStr(fullName, " ")

  GetLast = Right(fullName, Len(fullName) - space)
End Function

Sub DisplayLastFirst(firstName As String, lastName As String)
  MsgBox lastName & ", " & firstName
End Sub

