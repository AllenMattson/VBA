Attribute VB_Name = "Module5"
Option Explicit


Sub SendMsoMail(ByVal strRecipient As String)
' use MailEnvelope property of the Worksheet
' to return the msoEnvelope object
   
   ActiveWorkbook.EnvelopeVisible = True
   
   With ActiveSheet.MailEnvelope
   
        .Introduction = "Please see the list of  " & _
                    "employees who are to receive a bonus."
        With .Item
          ' Make sure the e-mail format is HTML
          .BodyFormat = olFormatHTML
          ' Add the recipient name
          .Recipients.Add strRecipient
          ' Add the subject
          .Subject = "Employee Bonuses"
          ' Send Mail
          .Send
        End With
   End With
End Sub





