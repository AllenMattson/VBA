Attribute VB_Name = "ApplicationTracker"
Sub btnNewApp_Click()
AllenJobsTrackerForm.Show vbModeless
End Sub
Sub btnClearTable_Click()
'backup progress into new worksheet and clear table
BackupMyWorkbook
Range("A4:H800").Cells.Clear
End Sub
Sub deletblankrows()
Dim LastRow As Long
LastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
For i = 1 To LastRow
    If Cells(i, 1) = "" Then
        Cells(i, 1).EntireRow.Delete
    End If
Next
End Sub

'Note: you may have to run the macro several times to delete all the blank rows!

'We can also use Autofilter to remove the blank rows:

Sub myautofilter()
Range(Selection, Selection.End(xlDown)).Select
    
    Selection.AutoFilter
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.AutoFilter
    ActiveSheet.Range("A:A").AutoFilter Field:=1, Criteria1:="<>"
    Selection.Copy
    Range("B1").Select
    ActiveSheet.Paste
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("B1").Select
    ActiveWorkbook.Save
End Sub

'To extract the required based on a keyword like  "Administration"  you can use the following VBA code:

Sub extractDatatoNeighboringColumn()
Dim LastRow As Long
LastRow = ActiveSheet.Range("A" & Rows.Count).End(xlUp).Row
For i = 1 To LastRow
    If InStr(Cells(i, 1).value, "Administration") Then
        Cells(i, 2).value = Cells(i, 1).value
    End If
Next
End Sub

Sub FindJobs()
Application.DisplayAlerts = False
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
    If ws.Name <> "Sheet1" Then ws.Delete
Next ws
Application.DisplayAlerts = True
Sheets.Add
'we define the essential variables
Dim ie As Object
Dim form As Variant, Button As Variant
Dim myjobtype As String, myexperience As String, mycity As String
'add the "Microsoft Internet Controls" reference in your VBA Project indirectly
Set ie = CreateObject("InternetExplorer.Application")

'more variables for the inputboxes - makes our automation program user friendly
myjobtype = InputBox("Enter type of job, eg. sales, administration")
myexperience = InputBox("enter your no of years experience, for example, 3")
mycity = InputBox("Enter the city where you wish to work")

With ie

.Visible = True
.navigate ("http://www.monsterindia.com")

' we ensure that the web page downloads completely before we fill the form automatically
While ie.ReadyState <> 4
DoEvents
Wend

'assigning the vinput variables to the html elements of the form
ie.Document.getelementsbyname("fts").Item.innertext = myjobtype
ie.Document.getelementsbyname("exp").Item(0).value = myexperience
ie.Document.getelementsbyname("lmy").Item.innertext = mycity
' accessing the button via the form
Set form = ie.Document.getElementsbytagname("form")

Set Button = form(0).onsubmit
form(0).submit

' again ensuring that the web page loads completely before we start scraping data
Do While ie.busy: DoEvents: Loop

Set TDelements = .Document.getElementsbytagname("td")
r = 0
c = 0

For Each TDelement In TDelements
ActiveSheet.Range("A1").Offset(r, c).value = TDelement.innertext
r = r + 1
Next

End With

' cleaning up memory
Set ie = Nothing
End Sub
Public Sub SaveMyJobDescription()

'Don't Forget to Add the Word Object Library in the Tools - References

' Call SetCRSDetails

Dim oCRSTemplate As String
'oCRSTemplate = gCRSPath & gCRSFileName

oCRSTemplate = "C:\Users\Allen\Desktop\JobApplicationTracker\Job Description\Template.docx"

Dim FilePicker As FileDialog
Dim objWord As Object
Dim objDocument As Object
Dim objWordAlreadyRunning As Boolean

objWordAlreadyRunning = False

On Error Resume Next
Set objWord = GetObject(, "Word.Application")
If Err.Number Then
Err.Clear
Set objWord = CreateObject("Word.Application")
If Err.Number Then
MsgBox "Can't open Word."
Set objDocument = Nothing
Set objWord = Nothing
Exit Sub
End If
Else
objWordAlreadyRunning = True
End If

objWord.Visible = True
objWord.Activate
objWord.Documents.Open (oCRSTemplate)
objWord.FileDialog(FileDialogType:=msoDialogSaveAs).Show
Set FilePicker = objWord.FileDialog(FileDialogType:=msoFileDialogSaveAs).Show

End Sub
Public Sub RunAfterJobDescription_FinalizeCells(Source As String)
Dim wb As Workbook: Set wb = ThisWorkbook
Dim ws As Worksheet: Set ws = wb.Sheets("Sheet1")

Dim lr As Integer, LC As Integer

ws.Select

'Fill in Row
lr = Cells(Rows.Count, 2).End(xlUp).Row
LC = 8 'Cells(3, Columns.Count).End(xlToLeft).Column


'Job Source

'Insert job source
If Len(Source) = 0 Then
    Cells(lr, LC).Offset(0, -1).value = "No Source Given"
Else
    Cells(lr, LC).Offset(0, -1).value = Source
End If



        If lr >= 3 Then
            Dim MyJobDescription As String
            MyJobDescription = "C:\Users\Allen\Desktop\JobApplicationTracker\Job Descriptions\" & FileName & ".docx"
            ActiveSheet.Hyperlinks.Add Range("F" & lr), MyJobDescription, , , "Job Description"
            'fill in resume and cover letter blanks as non submitted
            If Range("D" & lr) = "" Then Range("D" & lr).value = "Not Supplied"
            If Range("E" & lr).value = "" Then Range("E" & lr).value = "Not Supplied"
            'If Cells(lr, LC).Offset(0, -1).value <> "" Then Cells(lr, LC).Offset(0, -1).value = cbosource.Text
        Else

            MyJobDescription = "C:\Users\Allen\Desktop\JobApplicationTracker\Job Descriptions\" & FileName & ".docx"
            ActiveSheet.Hyperlinks.Add Range("F" & lr), MyJobDescription, , , "Job Description"
            'fill in resume and cover letter blanks as non submitted
            If Range("D" & lr) = "" Then Range("D" & lr).value = "Not Supplied"
            If Range("E" & lr).value = "" Then Range("E" & lr).value = "Not Supplied"
            'If Cells(lr, LC).Offset(0, -1).value <> "" Then Cells(lr, LC).Offset(0, -1).value = cbosource.Text
        End If

Columns.AutoFit
End Sub
Sub BackupMyWorkbook()
ThisWorkbook.SaveCopyAs _
FileName:=ThisWorkbook.Path & "\Backups\" & _
Format(Now(), "yyyy-mm-dd hh mm AMPM") & " " & _
ThisWorkbook.Name
End Sub


