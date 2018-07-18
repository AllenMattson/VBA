Attribute VB_Name = "Module1"
Option Explicit
Dim NextTick

Sub StartClock()
    UpdateClock
End Sub

Sub StopClock()
'   Cancels the OnTime event (stops the clock)
    On Error Resume Next
    Application.OnTime NextTick, "UpdateClock", , False
End Sub

Sub cbClockType_Click()
'   Hides or unhids the clock
    With ThisWorkbook.Sheets("Clock")
        If .DrawingObjects("cbClockType").Value = xlOn Then
            .ChartObjects("ClockChart").Visible = True
        Else
            .ChartObjects("ClockChart").Visible = False
        End If
    End With
End Sub

Sub UpdateClock()
'   Updates the clock that's visible
    Const PI As Double = 3.14159265358979
    Dim Clock As Chart
    Set Clock = ThisWorkbook.Sheets("Clock").ChartObjects("ClockChart").Chart
    
    If Clock.Parent.Visible Then
'       ANALOG CLOCK
        Dim CurrentSeries As Series
        Dim s As Series
        Dim x(1 To 2) As Variant
        Dim v(1 To 2) As Variant
    
'       Hour hand
        Set CurrentSeries = Clock.SeriesCollection("HourHand")
        x(1) = 0
        x(2) = 0.5 * Sin((Hour(Time) + (Minute(Time) / 60)) * (2 * PI / 12))
        v(1) = 0
        v(2) = 0.5 * Cos((Hour(Time) + (Minute(Time) / 60)) * (2 * PI / 12))
        CurrentSeries.XValues = x
        CurrentSeries.Values = v
        
'       Minute hand
        Set CurrentSeries = Clock.SeriesCollection("MinuteHand")
        x(1) = 0
        x(2) = 0.8 * Sin((Minute(Time) + (Second(Time) / 60)) * (2 * PI / 60))
        v(1) = 0
        v(2) = 0.8 * Cos((Minute(Time) + (Second(Time) / 60)) * (2 * PI / 60))
        CurrentSeries.XValues = x
        CurrentSeries.Values = v
    
'       Second hand
        Set CurrentSeries = Clock.SeriesCollection("SecondHand")
        x(1) = 0
        x(2) = 0.85 * Sin(Second(Time) * (2 * PI / 60))
        v(1) = 0
        v(2) = 0.85 * Cos(Second(Time) * (2 * PI / 60))
        CurrentSeries.XValues = x
        CurrentSeries.Values = v
    Else
'       DIGITAL CLOCK
        ThisWorkbook.Sheets("Clock").Range("DigitalClock").Value = CDbl(Time)
    End If
    
'   Set up the next event one second from now
    NextTick = Now + TimeValue("00:00:01")
    Application.OnTime NextTick, "UpdateClock"
End Sub
