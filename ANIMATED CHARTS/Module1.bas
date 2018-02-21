Attribute VB_Name = "Module1"
Option Explicit

Public Example1IsRunning As Boolean
Public Example2IsRunning As Boolean
Public Example3IsRunning As Boolean
Public Example4IsRunning As Boolean

Sub SimpleAnimation()
    Dim i As Long
    Range("A1") = 0
    For i = 1 To 150
       DoEvents
       Range("A1") = Range("A1") + 0.035
       DoEvents
    Next i
    Range("A1") = 0
End Sub


Sub RunExample1()
    If Example1IsRunning Then
        Example1IsRunning = False
        End
    End If
    Example1IsRunning = True
    Do
        DoEvents
        ThisWorkbook.Worksheets("Example-1").Range("Base") = _
            ThisWorkbook.Worksheets("Example-1").Range("Base") + 0.25
        DoEvents
    Loop
StopIt:
End Sub

Sub RunExample2()
Attribute RunExample2.VB_ProcData.VB_Invoke_Func = " \n14"
    If Example2IsRunning Then
        Example2IsRunning = False
        End
    End If
    Example2IsRunning = True
    
    Dim Multiplier As Object, Increment As Double, i As Double
    Set Multiplier = ThisWorkbook.Sheets("Example-2").Range("Multiplier")
    Do
        DoEvents
        Increment = 0.05
        Select Case Multiplier.Value
            Case Is > 0
                For i = 1 To -1 Step Increment * -1
                    DoEvents
                    Multiplier.Value = Application.Round(i, 2)
                    DoEvents
                Next i
            Case Is < 0
                For i = -1 To 1 Step Increment
                    DoEvents
                    Multiplier.Value = Application.Round(i, 2)
                    DoEvents
                Next i
            Case Is = 0
                For i = 0 To 1 Step Increment
                    DoEvents
                    Multiplier.Value = Application.Round(i, 2)
                    DoEvents
                Next i
                For i = 1 To -1 Step Increment * -1
                    DoEvents
                    Multiplier.Value = Application.Round(i, 2)
                    DoEvents
                Next i
                For i = -1 To 0 Step Increment
                    DoEvents
                    Multiplier.Value = Application.Round(i, 2)
                    DoEvents
                Next i
        End Select
    Loop
End Sub

Sub RunExample3()
    If Example3IsRunning Then
        Example3IsRunning = False
        End
    End If
    Example3IsRunning = True
    Do
        DoEvents
        ThisWorkbook.Sheets("Example-3").Range("Inc") = ThisWorkbook.Sheets("Example-3").Range("Inc") + 0.025
        DoEvents
    Loop
End Sub

''''''''Example-4 below
Sub Rotate1()
    Dim i As Long
    If Example4IsRunning Then
        Example4IsRunning = False
        End
    End If
    Example4IsRunning = True
    With ThisWorkbook.Sheets("Example-4").ChartObjects(1).Chart
        For i = 0 To 360 Step 8
            .Rotation = i
            '.Elevation = i - 90
            DoEvents
        Next i
    End With
    Example4IsRunning = False
End Sub


Sub Rotate2()
    Dim i As Long
'   Elevation range: -90 to +90
    If Example4IsRunning Then
        Example4IsRunning = False
        End
    End If
    Example4IsRunning = True
    With ThisWorkbook.Sheets("Example-4").ChartObjects(1).Chart
        For i = -90 To 90 Step 2
            .Elevation = i
            Application.Wait (Now + 0.000002)
            DoEvents
        Next i
        .Elevation = 15
    End With
    Example4IsRunning = False
End Sub

Sub Rotate3()
    Dim i As Long
'   Perspective range = 0 to 100
    If Example4IsRunning Then
        Example4IsRunning = False
        End
    End If
    Example4IsRunning = True
    With ThisWorkbook.Sheets("Example-4").ChartObjects(1).Chart
        For i = 0 To 100 Step 1
            .Perspective = i
            DoEvents
        Next i
        .Perspective = 30
    End With
    Example4IsRunning = False
End Sub

