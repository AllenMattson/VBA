Attribute VB_Name = "OpenWorkbookInstallAddin"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'AUTHOR: ALLEN MATTSON
'DATE: 5/7/2018
'DESCRIPTION:
'   This module will test for the version of Excel to make sure it
'   is greater than 2003 (Addin is for: 2007,2010,2013,2016).
'   Than it makes sure the addin isn't already in the addin's collection.
'   The user will be prompted if they would like to install this addin.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim InstalledProperly As Boolean
Private Sub workbook_AddInInstall()
    InstalledProperly = True
End Sub

Private Sub Workbook_Open()
Dim ai As AddIn, NewAi As AddIn
Dim M As String
Dim Ans As Long

'Check Version
If Val(Application.Version) < 12 Then
    MsgBox "This works only with Excel 2007 or later"
    ThisWorkbook.Close
End If

'Was just installed using the Add-Ins dialog box?
If InstalledProperly Then Exit Sub

For Each ai In AddIns
    If ai.Name = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & ".xlam" Then 'Replace the .xlsm with .xlam Then
        If ai.Installed Then
             MsgBox ThisWorkbook.Name & " add-in has been installed", vbInformation, ThisWorkbook.Name
             Exit Sub
        End If
    End If
Next ai

'Addin Not found, prompt user to install
M = "Install JEM2018 addin?" & vbNewLine & "Yes - Install JEM2018 add-in. "
M = M & vbNewLine & "No - Open it, but don't install it. "
M = M & vbNewLine & "Cancel - Close this workbook, close the add-in. I'm outa here."
Ans = MsgBox(M, vbQuestion + vbYesNoCancel, ThisWorkbook.Name)

Select Case Ans
    Case vbYes
        Call SaveWorkbookToAddinsCollection
        'add it to the Addins collection and install
        'Set NewAi = Application.AddIns.Add(ThisWorkbook.FullName)
        'NewAi.Installed = True
        
    Case vbNo
        'no action
    Case vbCancel
        ThisWorkbook.Close
End Select
    
End Sub
Private Sub SaveWorkbookToAddinsCollection()
Application.DisplayAlerts = False
'Set path variables to the users environment
Dim UserPath As String: UserPath = "C:\Users\"
Dim AddInsPath As String: AddInsPath = "AppData\Roaming\Microsoft\AddIns\" & Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5) & ".xlam" 'Replace the .xlsm with .xlam
Dim ActiveUser As String: ActiveUser = Environ("username") & "\"
On Error GoTo ErrHandler
'If ThisWorkbook.IsAddin = False Then ThisWorkbook.IsAddin = True
ThisWorkbook.SaveAs Filename:=UserPath & ActiveUser & AddInsPath, FileFormat:=xlOpenXMLAddIn
ErrHandler:
If Err.Number <> 0 Then MsgBox Err.Number & vbNewLine & Err.Description & vbNewLine & "On line: " & Erl
On Error GoTo 0
Application.DisplayAlerts = True
End Sub

