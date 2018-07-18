Attribute VB_Name = "Module1"
Option Explicit

Sub SendAsPDF()
'   Uses early binding
'   Requires a reference to the Outlook Object Library
    Dim OutlookApp As Outlook.Application
    Dim MItem As Object
    Dim Recipient As String, Subj As String
    Dim Msg As String, Fname As String
            
'   Message details
    Recipient = "myboss@xrediyh.com"
    Subj = "Sales figures"
    Msg = "Hey boss, here's the PDF file you wanted."
    Msg = Msg & vbNewLine & vbNewLine & "-Frank"
    Fname = Application.DefaultFilePath & "\" & _
      ActiveWorkbook.Name & ".pdf"
   
'   Create the attachment
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=Fname
    
'   Create Outlook object
    Set OutlookApp = New Outlook.Application
    
'   Create Mail Item and send it
    Set MItem = OutlookApp.CreateItem(olMailItem)
    With MItem
      .To = Recipient
      .Subject = Subj
      .Body = Msg
      .Attachments.Add Fname
      .Save 'to Drafts folder
      '.Send
    End With
    Set OutlookApp = Nothing

'   Delete the file
    Kill Fname
End Sub

