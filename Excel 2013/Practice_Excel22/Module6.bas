Attribute VB_Name = "Module6"
Option Explicit

Sub SendBulkMail(EmailCol, BeginRow, EndRow, SubjCol, NameCol, AmountCol)
    Dim objOut As Outlook.Application
    Dim objMail As Outlook.MailItem
    Dim strEmail As String
    Dim strSubject As String
    Dim strBody As String
    Dim r As Integer

    On Error Resume Next

    Application.DisplayAlerts = False

    Set objOut = New Outlook.Application

    For r = BeginRow To EndRow
        Set objMail = objOut.CreateItem(olMailItem)
        strEmail = Cells(r, EmailCol)
        strSubject = Cells(r, SubjCol) & " reimbursement"

        strBody = "Dear " & Cells(r, NameCol).Text & ":" & _
                    vbCrLf & vbCrLf
        strBody = strBody & "We have approved your request for " & _
                   LCase(strSubject)
        strBody = strBody & " in the amount of " & Cells(r, _
                   AmountCol).Text & "."
        strBody = strBody & vbCrLf & "Please allow 3 business " & _
                    "days for this"
        strBody = strBody & " amount to appear on your bank statement."
        strBody = strBody & vbCrLf & vbCrLf & " Employee Services"

        With objMail
            .To = strEmail
            .Body = strBody
            .Subject = strSubject
            '.Display
            .Send
        End With
    Next
    Set objOut = Nothing
    Application.DisplayAlerts = True
End Sub

Sub Call_SendBulkMail()
     SendBulkMail EmailCol:=4, _
          BeginRow:=2, _
          EndRow:=5, _
          SubjCol:=2, _
          NameCol:=1, _
          AmountCol:=3
End Sub




