Attribute VB_Name = "DownloadOutlookAttachments"
Sub AddRefToOutlook()
    Const outlookRef As String = "C:\Program Files (x86)\Microsoft Office\Office14\MSOUTL.OLB"

    If Not RefExists(outlookRef, "Microsoft Outlook 14.0 Object Library") Then
        Application.VBE.ActiveVBProject.References.AddFromFile _
            outlookRef
    End If
End Sub
Function RefExists(refPath As String, refDescrip As String) As Boolean
'Returns true/false if a specified reference exists, based on LIKE comparison
' to reference.description.

Dim ref As Variant
Dim bExists As Boolean

'Assume the reference doesn't exist
bExists = False

For Each ref In Application.VBE.ActiveVBProject.References
    If ref.Description Like refDescrip Then
        RefExists = True
        Exit Function
    End If
Next
RefExists = bExists
End Function
Sub sumit()

    readMails

End Sub


Function readMails()
    Dim olApp As Object ' OUTLOOK.Application
    Dim olNamespace As Object ' OUTLOOK.Namespace
    Dim olItem As Object ' OUTLOOK.MailItem
    Dim olInbox  As Object ' OUTLOOK.MAPIFolder
    Dim olFolder As Object ' OUTLOOK.MAPIFolder
    Dim oMsg As Object ' OUTLOOK.MailItem
    
    Dim i As Integer
    Dim b As Integer
    Dim lngCol As Long
    Dim mainWB As Workbook
    Dim keyword
    Dim Path
    Dim Count
    Dim Atmt
    Dim f_random
    Dim Filename
    'Dim olInbox As inbo
    Set olApp = CreateObject("Outlook.Application") ' New OUTLOOK.Application
    Set olNamespace = olApp.GetNamespace("MAPI")
     
     Set mainWB = ActiveWorkbook
     
    Set olInbox = olNamespace.GetDefaultFolder(olApp.olfolderinbox) '(OUTLOOK.olFolderInbox)
    Dim oItems As Object 'OUTLOOK.Items
   Set oItems = olInbox.Items
    mainWB.Sheets("Main").Range("A:A").Clear
    mainWB.Sheets("Main").Range("B:B").Clear
    mainWB.Sheets("Main").Range("A1,B1").Interior.ColorIndex = 46
    Path = mainWB.Sheets("Main").Range("J5").value
    keyword = mainWB.Sheets("Main").Range("J3").value
    mainWB.Sheets("Main").Range("A1").value = "Number"
    mainWB.Sheets("Main").Range("B1").value = "Subject"
    mainWB.Sheets("Main").Range("A1,B1").Borders.value = 1
    
    
    
    'MsgBox olInbox.Items.Count
   Count = 2
    For i = 1 To oItems.Count
        If TypeName(oItems.Item(i)) = "MailItem" Then
            Set oMsg = oItems.Item(i)
             
             If InStr(1, oMsg.Subject, keyword, vbTextCompare) > 0 Then
             'MsgBox "asfsdfsdf"
                    'MsgBox oMsg.Subject
                    mainWB.Sheets("Main").Range("A" & Count).value = Count - 1
                    mainWB.Sheets("Main").Range("B" & Count).value = oMsg.Subject
                    For Each Atmt In oMsg.Attachments
                    f_random = Replace(Replace(Replace(Now, " ", ""), "/", ""), ":", "") & "_"
                    Filename = Path & f_random & Atmt.Filename
                    'MsgBox Filename
                    Atmt.SaveAsFile Filename
                    FnWait (1)
                    Next Atmt
                    Count = Count + 1
             End If
        End If
        
    Next
   

End Function
Function FnWait(intTime)

    Dim newHour
    Dim NewMinute
    Dim newSecond
    Dim waitTime
    

    newHour = Hour(Now())

    NewMinute = Minute(Now())

    newSecond = Second(Now()) + intTime

     waitTime = TimeSerial(newHour, NewMinute, newSecond)

 Application.Wait waitTime

End Function
