Attribute VB_Name = "CalendarModule"
Sub BasicCalendar()
    dateVariable = CalendarForm.GetDate
    If dateVariable <> 0 Then Range("H16") = dateVariable
End Sub


Sub AdvancedCalendar()
    dateVariable = CalendarForm.GetDate( _
        SelectedDate:=Range("H34").value, _
        FirstDayOfWeek:=Monday, _
        DateFontSize:=12, _
        TodayButton:=True, _
        OkayButton:=True, _
        ShowWeekNumbers:=True, _
        BackgroundColor:=RGB(243, 249, 251), _
        HeaderColor:=RGB(147, 205, 2221), _
        HeaderFontColor:=RGB(255, 255, 255), _
        SubHeaderColor:=RGB(223, 240, 245), _
        SubHeaderFontColor:=RGB(31, 78, 120), _
        DateColor:=RGB(243, 249, 251), _
        DateFontColor:=RGB(31, 78, 120), _
        TrailingMonthFontColor:=RGB(155, 194, 230), _
        DateHoverColor:=RGB(223, 240, 245), _
        DateSelectedColor:=RGB(202, 223, 242), _
        SaturdayFontColor:=RGB(0, 176, 240), _
        SundayFontColor:=RGB(0, 176, 240), _
        TodayFontColor:=RGB(0, 176, 80))
    If dateVariable <> 0 Then Range("H34") = dateVariable
End Sub


Sub AdvancedCalendar2()
    dateVariable = CalendarForm.GetDate( _
        SelectedDate:=Range("H61").value, _
        DateFontSize:=11, _
        TodayButton:=True, _
        BackgroundColor:=RGB(242, 248, 238), _
        HeaderColor:=RGB(84, 130, 53), _
        HeaderFontColor:=RGB(255, 255, 255), _
        SubHeaderColor:=RGB(226, 239, 218), _
        SubHeaderFontColor:=RGB(55, 86, 35), _
        DateColor:=RGB(242, 248, 238), _
        DateFontColor:=RGB(55, 86, 35), _
        SaturdayFontColor:=RGB(55, 86, 35), _
        SundayFontColor:=RGB(55, 86, 35), _
        TrailingMonthFontColor:=RGB(106, 163, 67), _
        DateHoverColor:=RGB(198, 224, 180), _
        DateSelectedColor:=RGB(169, 208, 142), _
        TodayFontColor:=RGB(255, 0, 0), _
        DateSpecialEffect:=fmSpecialEffectRaised)
    If dateVariable <> 0 Then Range("H61") = dateVariable
End Sub


