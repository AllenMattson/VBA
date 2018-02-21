Attribute VB_Name = "Module1"
Option Explicit

Function NOMIDDLE(n) As String
    Dim FirstName As String, LastName As String
    n = Application.WorksheetFunction.Trim(n)
    FirstName = Left(n, InStr(1, n, " "))
    LastName = Right(n, Len(n) - InStrRev(n, " "))
    NOMIDDLE = FirstName & LastName
End Function

