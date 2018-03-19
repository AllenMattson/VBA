Attribute VB_Name = "Module3"
Private mypassword
Sub SBDC_Data_Tab()
Set WB = Nothing: Set TemplateWB = Nothing
Set WB = ThisWorkbook
'Dim WB As Workbook: Set WB = ActiveWorkbook
'set variables
Dim Title  As String
Dim FirstName  As String
Dim Surname  As String
Dim Telephone  As String
Dim Email  As String
Dim Address  As String
Dim Suburb  As String
Dim State  As String
Dim Postcode  As String
Dim BusinessDuration  As String
Dim ANZICCODE  As String
Dim AboriginalBusiness  As String
Dim BusinessName  As String
Dim ABN  As String
Dim IndigenousInBusiness  As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CHECK FOR TEMPLATE PATH, IF IT EXISTS CONTINUE, ELSE PROMPT FOR PATH
Dim TemplatePath As String
WB.Activate:
Sheets("Sheet3").Select
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Len(Range("B4")) > 2 Then
    TemplatePath = Sheets("Sheet3").Cells(4, 2).Value
Else
    TemplatePath = InputBox("Please Insert Template File Path...", "Template Sheet Location", WB.path)
    'Sheets("Sheet3").Cells(1, 2).Value = TemplatePath
End If
If TemplatePath = "" Then
    MsgBox "Insert Path to Template Files"
    Exit Sub
End If
If TemplatePath = "" Then
    MsgBox "Insert Template Path": Exit Sub
End If
'insert template path into sheet3
Sheets("Sheet3").Activate: Range("B4").Value = TemplatePath
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Locate Template Sheet SBDC from cell 'b5' or prompt for the name
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim SBDC As String
If Len(Range("B5")) > 2 Then
    SBDC = Sheets("Sheet3").Cells(5, 2).Value
Else
    SBDC = InputBox("Enter SBDC Template workbook name", "SDBC", "SBDCreport.xlsx")
    Range("B5") = SBDC
End If
If SBDC = "" Then
    MsgBox "Insert SBDC template sheet name...", vbOKOnly, "Please try again"
    Exit Sub
End If

'On Error GoTo TempUnfound
Set TemplateWB = Workbooks.Open(TemplatePath & SBDC)
Sheets("Data").Select
Dim StartCell As Range: Set StartCell = Range("A90000").End(xlUp)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'Dim TemplateWB As Workbook
WB.Activate
Dim LR As Integer: LR = Range("A90000").End(xlUp).Row
Dim i As Integer
Sheets("Sheet1").Activate
For i = LR To 2 Step -1
If ActiveWorkbook.Name <> WB.Name Then WB.Activate 'TemplateWB.Active = True Then WB.Activate
Sheets("Sheet1").Activate
'check to see if this is correct template
If Cells(i, "F").Value <> "Business Local" Then GoTo NextROw


Title = Cells(i, 9).Value: Debug.Print Title
FirstName = Cells(i, 10).Value: Debug.Print FirstName
Surname = Cells(i, 11).Value: Debug.Print Surname
Telephone = Cells(i, 12).Value: Debug.Print Telephone
Email = Cells(i, 13).Value: Debug.Print Email
Address = Cells(i, 14).Value: Debug.Print Address
Suburb = Cells(i, 15).Value: Debug.Print Suburb
State = Cells(i, 16).Value: Debug.Print State
Postcode = Cells(i, 17).Value: Debug.Print Postcode
BusinessDuration = Cells(i, 18).Value: Debug.Print BusinessDuration
ANZICCODE = Cells(i, 19).Value: Debug.Print ANZICCODE
AboriginalBusiness = Cells(i, 20).Value: Debug.Print AboriginalBusiness
BusinessName = Cells(i, 21).Value: Debug.Print BusinessName
ABN = Cells(i, "V").Value: Debug.Print ABN
IndigenousInBusiness = Cells(i, "T").Value: Debug.Print IndigenousInBusiness



TemplateWB.Activate

Dim LC As Integer: LC = Cells(1, Columns.Count).End(xlToLeft).Column
Dim j As Integer
Dim sh As Worksheet
'dynamically remove protection to insert values
'protect cells again
PasswordBreaker
ThisWorkbook.Unprotect mypassword



For j = 1 To LC
'    If Cells.protected = True Then Cells.protected = False
    StartCell.Offset(0, j).Value = Title & " _" & j
    StartCell.Offset(0, j).Value = FirstName & " _" & j
    StartCell.Offset(0, j).Value = Surname & " _" & j
    StartCell.Offset(0, j).Value = Telephone & " _" & j
    StartCell.Offset(0, j).Value = Email & " _" & j
    StartCell.Offset(0, j).Value = Address & " _" & j
    StartCell.Offset(0, j).Value = Suburb & " _" & j
    StartCell.Offset(0, j).Value = State & " _" & j
    StartCell.Offset(0, j).Value = Postcode & " _" & j
    StartCell.Offset(0, j).Value = BusinessDuration & " _" & j
    StartCell.Offset(0, j).Value = ANZICCODE & " _" & j
    StartCell.Offset(0, j).Value = AboriginalBusiness & " _" & j
    StartCell.Offset(0, j).Value = BusinessName & " _" & j
    StartCell.Offset(0, j).Value = ABN & " _" & j
    StartCell.Offset(0, j).Value = IndigenousInBusiness & " _" & j
Next j

NextROw:
Next i
MsgBox "done"
'lock cells again
TemplateWB.Activate
For Each sh In ActiveWorkbook.Worksheets
    sh.Protect Password:=mypassword
Next sh
'Exit Sub
''''''''''''''''''''''

'''''''''''''''''''''''''
End Sub
Sub ProtectAll()
    Dim sh As Worksheet
'dynamically remove protection to insert values
'protect cells again
    PasswordBreaker
    ThisWorkbook.Unprotect mypassword

    For Each sh In ActiveWorkbook.Worksheets
        sh.Protect Password:=mypassword
    Next sh
    
End Sub
Function PasswordBreaker()
    'Breaks worksheet password protection.
    Dim i As Integer, j As Integer, k As Integer
    Dim l As Integer, m As Integer, n As Integer
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i4 As Integer, i5 As Integer, i6 As Integer
    On Error Resume Next
    For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
    For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
    For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
    For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
    ActiveSheet.Unprotect Chr(i) & Chr(j) & Chr(k) & _
        Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
    If ActiveSheet.ProtectContents = False Then
            mypassword = Chr(i) & Chr(j) & _
            Chr(k) & Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & _
            Chr(i3) & Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
         Exit Function
    End If
    Next: Next: Next: Next: Next: Next
    Next: Next: Next: Next: Next: Next
End Function
