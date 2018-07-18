Attribute VB_Name = "Module1"
Sub Macro86()

'Step 1:  Declare your variables
    Dim OLApp As Outlook.Application
    Dim OLMail As Object
    Dim MyCell As Range
    Dim MyContacts As Range
    
'Step 2:  Define the range to loop through
    Set MyContacts = Sheets("Contact List").Range("H2:H21")
    
'Step 3:  Open Outlook
    Set OLApp = New Outlook.Application
    Set OLMail = OLApp.CreateItem(0)
    OLApp.Session.Logon
    
'Step 4:  Add each address in the contact list
    With OLMail
        .BCC = ""
          For Each MyCell In MyContacts
            .BCC = .BCC & MyCell.Value & ";"
          Next MyCell
        .Subject = "Chapter 11 Sample Email"
        .Body = "Sample file is attached"
        .Attachments.Add ActiveWorkbook.FullName
        .Display  'Change to .Send to send without reviewing
    End With
    
'Step 5:  Memory cleanup
    Set OLMail = Nothing
    Set OLApp = Nothing
    
    
End Sub

