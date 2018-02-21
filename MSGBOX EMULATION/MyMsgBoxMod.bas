Attribute VB_Name = "MyMsgBoxMod"
Option Explicit

#If VBA7 And Win64 Then
    Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Public UserClick As Integer
Public Prompt1 As String
Public Buttons1 As Integer
Public Title1 As String

Function MyMsgBox(ByVal Prompt As String, Optional ByVal Buttons As Integer, Optional ByVal Title As String) As Integer
'   Emulates VBA's MsgBox function
'   Does not support the HelpFile or Context arguments
    
'   Pass the function arguments to the UserForm
'   By using public variables
    Prompt1 = Prompt
    Buttons1 = Buttons
    Title1 = Title
    
'   Show the form
'   The work is done in the Initialize event-handler
    With MyMsgBoxForm
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
'   Get the return value
    MyMsgBox = UserClick
End Function

Sub TestMyMsgBox()
    Dim Prompt As String
    Dim Buttons As Integer
    Dim Title As String
    
'   Tests the MyMsgbox Function
'   Can be deleted from this module
    Prompt = Range("B1")
    Buttons = Range("B2")
    Title = Range("B3")
    Range("B4") = MyMsgBox(Prompt, Buttons, Title)
End Sub

Sub TestMsgBox()
    Dim Prompt As String
    Dim Buttons As Integer
    Dim Title As String

'   Tests the standard Msgbox Function
'   Can be deleted from this module
    Prompt = Range("B1")
    Buttons = Range("B2")
    Title = Range("B3")
    Range("B4") = MsgBox(Prompt, Buttons, Title)
End Sub

Sub MyMsgBoxTest()
    Dim Prompt, Buttons, Title, Ans
    Prompt = "You have chosen to save this workbook" & vbCrLf
    Prompt = Prompt & "on a drive that is not available to" & vbCrLf
    Prompt = Prompt & "all employees." & vbCrLf & vbCrLf
    Prompt = Prompt & "OK to continue?"
    Buttons = vbQuestion + vbYesNo
    Title = "We have a problem"
    Ans = MyMsgBox(Prompt, Buttons, Title)
End Sub
