Attribute VB_Name = "Module1"
Option Explicit

Function Commission(Sales)
    Const Tier1 = 0.08
    Const Tier2 = 0.105
    Const Tier3 = 0.12
    Const Tier4 = 0.14
'   Calculates sales commissions
    Select Case Sales
       Case 0 To 9999.99: Commission = Sales * Tier1
       Case 10000 To 19999.99: Commission = Sales * Tier2
       Case 20000 To 39999.99: Commission = Sales * Tier3
       Case Is >= 40000: Commission = Sales * Tier4
    End Select
End Function

Function Commission2(Sales, Years) As Single
'    Calculates sales commissions based on
'    years in service
     Const Tier1 = 0.08
     Const Tier2 = 0.105
     Const Tier3 = 0.12
     Const Tier4 = 0.14
     Select Case Sales
        Case 0 To 9999.99: Commission2 = Sales * Tier1
        Case 10000 To 19999.99: Commission2 = Sales * Tier2
        Case 20000 To 39999.99: Commission2 = Sales * Tier3
        Case Is >= 40000: Commission2 = Sales * Tier4
     End Select
     Commission2 = Commission2 + (Commission2 * Years / 100)
End Function


Sub CalcComm()
    Dim Sales As Double
    Dim Message As String
    Dim Ans As Integer
    
'   Prompt for sales amount
    Sales = Val(InputBox("Enter Sales:", _
     "Sales Commission Calculator"))

'   Build the message
    Message = "Sales Amount:" & vbTab & Format(Sales, "$#,##0")
    Message = Message & vbCrLf & "Commission:" & vbTab
    Message = Message & Format(Commission(Sales), "$#,##0")
    Message = Message & vbCrLf & vbCrLf & "Another?"

'   Display the result and prompt for another
    Ans = MsgBox(Message, vbYesNo, "Sales Commission Calculator")
    If Ans = vbYes Then CalcComm
End Sub


