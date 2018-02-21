Attribute VB_Name = "Conditional"
' declare a conditional compiler constant
#Const verPolish = True

Sub WhatDay()
    Dim dayNr As Integer

    #If verPolish = True Then
        dayNr = Weekday(InputBox("Wpisz date, np. 06/04/2013"))
        MsgBox "To bedzie " & DayOfWeek(dayNr) & "."
    #Else
        WeekdayName
    #End If
End Sub

Function DayOfWeek(dayNr As Integer) As String
    DayOfWeek = Choose(dayNr, "niedziela", "poniedzialek", "wtorek", _
    "sroda", "czwartek", "piatek", "sobota")
End Function

Function WeekdayName() As String
    Select Case Weekday(InputBox("Enter date, e.g., 06/04/2013"))
        Case 1
            WeekdayName = "Sunday"
        Case 2
            WeekdayName = "Monday"
        Case 3
            WeekdayName = "Tuesday"
        Case 4
            WeekdayName = "Wednesday"
        Case 5
            WeekdayName = "Thursday"
        Case 6
            WeekdayName = "Friday"
        Case 7
            WeekdayName = "Saturday"
    End Select
    MsgBox "It will be " & WeekdayName & "."
End Function


