Attribute VB_Name = "Module1"
Option Explicit

Sub SendEmail()
  'Uses early binding
  'Requires a reference to the Outlook Object Library
  Dim OutlookApp As Outlook.Application
  Dim MItem As Outlook.MailItem
  Dim cell As Range
  Dim Subj As String
  Dim EmailAddr As String
  Dim Recipient As String
  Dim Bonus As String
  Dim Msg As String
  
  'Create Outlook object
  Set OutlookApp = New Outlook.Application
  
  'Loop through the rows
  For Each cell In Columns("B").Cells.SpecialCells(xlCellTypeConstants)
    If cell.Value Like "*@*" Then
      'Get the data
      Subj = "Your Annual Bonus"
      Recipient = cell.Offset(0, -1).Value
      EmailAddr = cell.Value
      Bonus = Format(cell.Offset(0, 1).Value, "$0,000.")
            
     'Compose message
      Msg = "Dear " & Recipient & vbCrLf & vbCrLf
      Msg = Msg & "I am pleased to inform you that your annual bonus is "
      Msg = Msg & Bonus & vbCrLf & vbCrLf
      Msg = Msg & "William Rose" & vbCrLf
      Msg = Msg & "President"
    
      'Create Mail Item and send it
      Set MItem = OutlookApp.CreateItem(olMailItem)
      With MItem
        .To = EmailAddr
        .Subject = Subj
        .Body = Msg
        '.Send
        .Save 'to Drafts folder
      End With
    End If
  Next
  Set OutlookApp = Nothing
End Sub

