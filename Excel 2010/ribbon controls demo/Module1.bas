Attribute VB_Name = "Module1"
'Callback for Label1 getLabel
Sub getLabel1(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Hello " & Application.UserName
End Sub

'Callback for Label2 getLabel
Sub getLabel2(control As IRibbonControl, ByRef returnedVal)
    returnedVal = "Today is " & Date
End Sub

'Callback for EditBox1 onChange
Sub EditBox1_Change(control As IRibbonControl, text As String)
    Dim squareRoot As Double
    On Error Resume Next
    squareRoot = Sqr(text)
    If Err.Number = 0 Then
        MsgBox "The square root of " & text & " is: " & squareRoot
    Else
        MsgBox "Enter a positive number.", vbCritical
    End If
End Sub

'Callback for Button1 onAction
Sub ShowCalculator(control As IRibbonControl)
    On Error Resume Next
    Shell "calc.exe", vbNormalFocus
    If Err.Number <> 0 Then MsgBox "Can't start calc.exe"
End Sub


'Callback for ToggleButton1 getPressed
Sub ToggleButton1_Click(control As IRibbonControl, ByRef returnedVal)
    MsgBox "Toggle value: " & returnedVal
End Sub


'Callback for Checkbox1 onAction
Sub Checkbox1_Change(control As IRibbonControl, pressed As Boolean)
    MsgBox "Checkbox value: " & pressed
End Sub

'Callback for Combo1 onChange
Sub Combo1_Change(control As IRibbonControl, text As String)
    MsgBox text
End Sub

'Callback for Gallery1 onAction
Sub MonthSelected(control As IRibbonControl, _
   id As String, index As Integer)
    MsgBox "You selected " & id
End Sub

'Callback for Today onAction
Sub ShowToday(control As IRibbonControl)
    MsgBox "Today is " & Date
End Sub

'Callback for gallery2 onAction
Sub OnAction(control As IRibbonControl, id As String, index As Integer)
    MsgBox "You clicked Banjo Photo #" & index + 1
End Sub

