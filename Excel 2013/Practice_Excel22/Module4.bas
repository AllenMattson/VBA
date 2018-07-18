Attribute VB_Name = "Module4"
Option Explicit

Sub SendMailNow()
    Dim strEAddress As String

    On Error GoTo ErrorHandler

    strEAddress = InputBox("Enter e-mail address", _
                "Recipient's E-mail Address ")

    If IsNull(Application.MailSession) Then
        Application.MailLogon
    End If

    ActiveWorkbook.SendMail Recipients:=strEAddress, Subject:="Test Mail"

    Application.MailLogoff
    Exit Sub

ErrorHandler:
    MsgBox "Some error occurred while sending e-mail."
End Sub



