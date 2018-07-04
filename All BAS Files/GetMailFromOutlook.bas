Attribute VB_Name = "GetMailFromOutlook"
Option Explicit

Private lRow As Long, x As Date, oWS As Worksheet

Sub GetFromInbox()
ActiveSheet.Cells.Clear
Cells(1, 1).Value = "Subject"
Cells(1, 2).Value = "Received Time"
Cells(1, 3).Value = "Sender Name"

    Const olFolderInbox = 6
    Dim olApp As Object, olNs As Object
    Dim oRootFldr As Object ' Root folder to start
    Dim lCalcMode As Long

    Set olApp = CreateObject("Outlook.Application")
    Set olNs = olApp.GetNamespace("MAPI")
    Set oRootFldr = olNs.GetDefaultFolder(olFolderInbox).Folders("Job Websites")
    Set oWS = ActiveSheet

    x = Date
    lRow = 1
    lCalcMode = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    GetFromFolder oRootFldr
    Application.ScreenUpdating = True
    Application.Calculation = lCalcMode

    Set oWS = Nothing
    Set oRootFldr = Nothing
    Set olNs = Nothing
    Set olApp = Nothing
End Sub

Private Sub GetFromFolder(oFldr As Object)
    Dim oItem As Object, oSubFldr As Object
lRow = 2
    ' Process all mail items in this folder
    For Each oItem In oFldr.Items
        If TypeName(oItem) = "MailItem" Then
            With oItem
               ' If InStr(1, .Subject, "TechFetch", vbTextCompare) > 0 And DateDiff("d", .ReceivedTime, x) = 0 Then
                    oWS.Cells(lRow, 1).Value = .Subject
                    oWS.Cells(lRow, 2).Value = .ReceivedTime
                    oWS.Cells(lRow, 3).Value = .SenderName
                    lRow = lRow + 1
               ' End If
            End With
        End If
    Next

    ' Recurse all Subfolders
    For Each oSubFldr In oFldr.Folders
        GetFromFolder oSubFldr
    Next
End Sub
