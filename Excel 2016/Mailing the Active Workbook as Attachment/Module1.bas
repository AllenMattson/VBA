Attribute VB_Name = "Module1"
Sub Macro85()

'Step 1:  Declare your variables
    Dim OLApp As Outlook.Application
    Dim OLMail As Object
    
    
'Step 2:  Open Outlook start a new mail item
    Set OLApp = New Outlook.Application
    Set OLMail = OLApp.CreateItem(0)
    OLApp.Session.Logon
    
    
'Step 3:  Build your mail item and send
    With OLMail
    .To = "admin@datapigtechnologies.com; mike@datapigtechnologies.com"
    .CC = ""
    .BCC = ""
    .Subject = "This is the Subject line"
    .Body = "Hi there"
    .Attachments.Add ActiveWorkbook.FullName
    .Display  'Change to .Send to send without reviewing
    End With
    
    
'Step 4:  Memory cleanup
    Set OLMail = Nothing
    Set OLApp = Nothing

End Sub

