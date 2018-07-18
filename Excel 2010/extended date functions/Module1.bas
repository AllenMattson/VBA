Attribute VB_Name = "Module1"
Option Explicit

Function XDATE(y, m, d, Optional fmt As String) As String
Attribute XDATE.VB_Description = "Returns a date for any year between 0100 and 9999. fmt is an optional date formatting string."
Attribute XDATE.VB_HelpID = 200
Attribute XDATE.VB_ProcData.VB_Invoke_Func = " \n2"
    If IsMissing(fmt) Then fmt = "Short Date"
    XDATE = Format(DateSerial(y, m, d), fmt)
End Function

Function XDATEADD(xdate1, days, Optional fmt As String) As String
Attribute XDATEADD.VB_Description = "Returns a date, incremented by a specified number of days. fmt is an optional date formatting string."
Attribute XDATEADD.VB_HelpID = 300
Attribute XDATEADD.VB_ProcData.VB_Invoke_Func = " \n2"
    Dim TempDate As Date
    If IsMissing(fmt) Then fmt = "Short Date"
    xdate1 = RemoveDay(xdate1)
    TempDate = DateValue(xdate1)
    XDATEADD = Format(TempDate + days, fmt)
End Function

Function XDATEDIF(xdate1, xdate2) As Long
Attribute XDATEDIF.VB_Description = "Returns the number of days between date1 and date2 (date1-date2)."
Attribute XDATEDIF.VB_HelpID = 400
Attribute XDATEDIF.VB_ProcData.VB_Invoke_Func = " \n2"
    xdate1 = RemoveDay(xdate1)
    xdate2 = RemoveDay(xdate2)
    XDATEDIF = DateSerial(Year(xdate1), Month(xdate1), Day(xdate1)) - DateSerial(Year(xdate2), Month(xdate2), Day(xdate2))
End Function

Function XDATEYEARDIF(xdate1, xdate2) As Long
Attribute XDATEYEARDIF.VB_Description = "Returns the number of full years between date1 and date2 (date1-date2). Useful for calculating ages."
Attribute XDATEYEARDIF.VB_HelpID = 500
Attribute XDATEYEARDIF.VB_ProcData.VB_Invoke_Func = " \n2"
    Dim YearDiff As Long
    xdate1 = RemoveDay(xdate1)
    xdate2 = RemoveDay(xdate2)
    YearDiff = Year(xdate2) - Year(xdate1)
    If DateSerial(Year(xdate1), Month(xdate2), Day(xdate2)) < CDate(xdate1) Then YearDiff = YearDiff - 1
    XDATEYEARDIF = YearDiff
End Function

Function XDATEYEAR(xdate1)
Attribute XDATEYEAR.VB_Description = "Returns the year for a date."
Attribute XDATEYEAR.VB_HelpID = 600
Attribute XDATEYEAR.VB_ProcData.VB_Invoke_Func = " \n2"
    xdate1 = RemoveDay(xdate1)
    XDATEYEAR = Year(DateValue(xdate1))
End Function

Function XDATEMONTH(xdate1)
Attribute XDATEMONTH.VB_Description = "Returns the month for a date."
Attribute XDATEMONTH.VB_HelpID = 700
Attribute XDATEMONTH.VB_ProcData.VB_Invoke_Func = " \n2"
    xdate1 = RemoveDay(xdate1)
    XDATEMONTH = Month(DateValue(xdate1))
End Function

Function XDATEDAY(xdate1)
Attribute XDATEDAY.VB_Description = "Returns the day for a date."
Attribute XDATEDAY.VB_HelpID = 800
Attribute XDATEDAY.VB_ProcData.VB_Invoke_Func = " \n2"
    xdate1 = RemoveDay(xdate1)
    XDATEDAY = Day(DateValue(xdate1))
End Function

Function XDATEDOW(xdate1)
Attribute XDATEDOW.VB_Description = "Returns an integer corresponding to the weekday for a date (1=Sunday)."
Attribute XDATEDOW.VB_HelpID = 900
Attribute XDATEDOW.VB_ProcData.VB_Invoke_Func = " \n2"
    xdate1 = RemoveDay(xdate1)
    XDATEDOW = Weekday(xdate1)
End Function

Private Function RemoveDay(xdate1)
'   Remove day of week from string
    Dim i As Integer
    Dim Temp As String
    Temp = xdate1
    For i = 0 To 6 'Unabbreviated day names
        Temp = Application.Substitute(Temp, Format(DateSerial(1900, 1, 0), "dddd"), "")
    Next i
    For i = 0 To 6 'Abbreviated day names
        Temp = Application.Substitute(Temp, Format(DateSerial(1900, 1, 0), "ddd"), "")
    Next i
    RemoveDay = Temp
End Function

Sub SetMacroOptions()
'   Add descriptions, and put in the Date & Time function category
    On Error Resume Next
    With Application
        .MacroOptions macro:="XDATE", Description:="Returns a date for any year between 0100 and 9999. fmt is an optional date formatting string.", Category:=2
        .MacroOptions macro:="XDATEADD", Description:="Returns a date, incremented by a specified number of days. fmt is an optional date formatting string.", Category:=2
        .MacroOptions macro:="XDATEDIF", Description:="Returns the number of days between date1 and date2 (date1-date2).", Category:=2
        .MacroOptions macro:="XDATEYEARDIF", Description:="Returns the number of full years between date1 and date2 (date1-date2). Useful for calculating ages."
        .MacroOptions macro:="XDATEYEAR", Description:="Returns the year for a date."
        .MacroOptions macro:="XDATEMONTH", Description:="Returns the month for a date."
        .MacroOptions macro:="XDATEDAY", Description:="Returns the day for a date."
        .MacroOptions macro:="XDATEDOW", Description:="Returns an integer corresponding to the weekday for a date (1=Sunday)."
    End With
End Sub


