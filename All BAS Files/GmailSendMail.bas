Attribute VB_Name = "GmailSendMail"
Sub CDO_Mail_Small_Text()
    Dim iMsg As Object
    Dim iConf As Object
    Dim strbody As String
    Dim Flds As Variant
    Dim PDFfileName As String
    
    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
On Error GoTo ErrHandle
    iConf.Load -1    ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "mattson.allen@gmail.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "34553455t"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
        .Update
    End With

    'PDFfileName = Cells(28, 3)   'This needs to be the complete path and filename if it is used, I kept that info in C28.

    strbody = "Automated email" & vbNewLine & vbNewLine & _
              "Info Here" & vbNewLine & vbNewLine & _
              "Your Name Here"

    With iMsg
        Set .Configuration = iConf
        .To = "mattson.allen@gmail.com"   'this can be any email address
        .CC = ""
        .BCC = ""
        .From = "mattson.allen@gmail.com"    'MUST BE THE SAME AS USED ABOVE
        .Subject = "Automated Test" 'Cells(4, 6) & " " & Cells(4, 8)      'I had the Month and Year in these two cells, remove and change
        .TextBody = strbody
        '.AddAttachment     'PDFfileName  - this was my calendar
        .Send
    End With
MsgBox "Message Sent!!"
ErrHandle:
If Err.Number <> 0 Then
    Sheets.Add
    Cells(1, 1) = "Error Number"
    Cells(2, 1) = Err.Number
    Cells(1, 2) = "Error Description"
    Cells(2, 2) = Err.Description
    Cells(1, 3) = "Help Context"
    Cells(2, 3) = Err.HelpContext
    Cells(1, 4) = "Help File"
    Cells(2, 4) = Err.HelpFile
    Columns.AutoFit
    MsgBox "Error Number: " & Err.Number & vbNewLine & vbNewLine & Err.Description
End If
    
'Set the NewMail Variable to Nothing
    Set iMsg = Nothing
    Set iConf = Nothing
End Sub
