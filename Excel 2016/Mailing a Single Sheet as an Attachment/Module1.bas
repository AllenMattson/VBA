Attribute VB_Name = "Module1"
Sub Macro87()

'Step 1:  Declare your variables
    Dim OLApp As Outlook.Application
    Dim OLMail As Object
    
    
'Step 2:  Copy Workheet, paste to new workbook, and save it
    Sheets("Revenue Table").Copy
    ActiveWorkbook.SaveAs ThisWorkbook.Path & "\TempRangeForEmail.xlsx"
    
    
'Step 3:  Open Outlook start a new mail item
    Set OLApp = New Outlook.Application
    Set OLMail = OLApp.CreateItem(0)
    OLApp.Session.Logon
    
    
'Step 4:  Build your mail item and send
    With OLMail
    .To = "admin@datapigtechnologies.com; mike@datapigtechnologies.com"
    .CC = ""
    .BCC = ""
    .Subject = "This is the Subject line"
    .Body = "New Workbook Attached"
    .Attachments.Add (ThisWorkbook.Path & "\TempRangeForEmail.xlsx")
    .Display  'Change to .Send to send without reviewing
    End With
    
    
'Step 5:  Delete the temporary Excel file
    ActiveWorkbook.Close SaveChanges:=True
    Kill ThisWorkbook.Path & "\TempRangeForEmail.xlsx"
    
    
'Step 6:  Memory cleanup
    Set OLMail = Nothing
    Set OLApp = Nothing
    
End Sub

