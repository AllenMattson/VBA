Attribute VB_Name = "ExportChildWorkBooks"
Option Explicit

Private MyDate As Date
Private Jobnumber As Single
Private Topic As String
Private LeadSource As String
Private STATUS As String
Private Program As String
Private BusAdviser As String
Private ClientID As String
Private FirstName As String
Private Surname As String
Private Telephone As String
Private Email As String
Private Address As String
Private Suburb As String
Private State As String
Private Postcode As Integer
Private BusinessDuration As Integer
Private ANZICCODE As String
Private AboriginalBusiness As String
Private BusinessName As String
Private ABN As String
Private CURRENTNumberofEmployees As Integer
Private WomenInBusiness As String
Private IndigenousInBusiness As String
Private FamilyInBusiness As String
Private Homebasedbusiness As String
Private Fundingavenuesandfinancialanalysis As String
Private Buildingyourbusiness As String
Private Makingthemostofyourtalentandteam As String
Private Managementcapabilities As String
Private Digitalengagementimplementation As String
Private Tourismready As String
Private LegalNameofBusiness As String
Private ABNforTHISbusiness As String
Private SmallBusinessIntenderhasNOABN As String
Private Consenttobesurveyed As String
Private firstentrytoproject As String
Private businessdiagnosticcompletedthisquarter As String
Private UniqueBusiness As String
Private SUST As String
Private JobStartDate As Date
Private JobEndDate As Date
Private AgginreasedTO As Single
Private valuenewpositions As Single
Private Newpositions As Integer
Private valueCapInv As Single
Private HOURSPROVIDED As Integer
Private FollowUp As Integer
Private Referredforwards As String
Private InvDate As Date
Private Inv As Integer
Private PaidDate As Date
Private Amount As Single
Private Notes As String

Public ChildSheetLocation As String

Sub Child_Workbooks()
Dim s As Worksheet, sh As Worksheet
Dim Master_WB As Workbook: Set Master_WB = ThisWorkbook

Application.ScreenUpdating = False
'delete old sheets if there is any
Dim OldSH As Worksheet
For Each OldSH In ThisWorkbook.Sheets
    If OldSH.Name <> "Sheet1" Then
        If OldSH.Name <> "Sheet2" Then
            If OldSH.Name <> "Sheet3" Then
                Application.DisplayAlerts = False
                OldSH.Delete
                Application.DisplayAlerts = True
            End If
        End If
    End If
Next OldSH
'Make New Folder
Dim MyPathFolder As String: MyPathFolder = InputBox("Input file path to upload sheets to...", "Upload Advisor Data", ThisWorkbook.Path & "\")
If MyPathFolder = "" Then Exit Sub
'make sure there is a backslash to end file path
    If Right(MyPathFolder, 1) <> "\" Then
        MyPathFolder = MyPathFolder & "\"
    End If
'Identify Location of Child Sheets
ChildSheetLocation = MyPathFolder


ClearFolder (MyPathFolder)
Master_WB.Activate
Sheets("Sheet2").Activate
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Dim PromptForImport As String: PromptForImport = MsgBox("Would you like to import the new advisor data?", vbYesNo, "Import New Advisor Data before uploading")
'If PromptForImport = vbYes Then
    'combine data into new master sheet
'    ImportCSVfiles.Merge_CSV_Files
'    Clear_All_CSV_files_In_Folder
'Else
    'remove csv files if any are there
'    Clear_All_CSV_files_In_Folder
'End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Adv_BookName As String
'Loop Advisor Names and create worksheets in new folder
'save as csv to open with google and excel

Dim Adv_RNG As Range: Set Adv_RNG = Sheets("Sheet2").Range("E2:E" & Rows.Count)
Dim cell As Range
Dim LastBArow As Integer, BA As Integer
For Each cell In Adv_RNG
    If cell.Value <> "" Then
        Adv_BookName = cell.Value
        Sheets("Sheet1").Cells.Copy
        Sheets.Add
        Cells(1, 1).PasteSpecial xlPasteAll
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'remove secret data
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        LastBArow = Cells(Rows.Count, 6).End(xlUp).Row
        For BA = LastBArow To 2 Step -1
            If Cells(BA, 6).Value <> Adv_BookName Then Rows(BA).Delete
        Next BA
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'move to new workbook save as BA name
        ActiveSheet.Move
        ActiveWorkbook.SaveAs FileName:=MyPathFolder & Adv_BookName & ".csv", FileFormat:=xlCSV '
        ActiveWorkbook.Close True
        Master_WB.Activate
        Sheets("Sheet1").Activate
        'Create xls docs
        ExportChildWorkBooks.LoopThroughFilesChangeCSVtoXLS (MyPathFolder)
        'remove csv files
        'advisors will work off the xls files
        ExportChildWorkBooks.Clear_All_CSV_files_In_Folder
    End If
Next cell





Application.ScreenUpdating = True

End Sub
Sub LoopThroughFilesChangeCSVtoXLS(MyPathFolder As String)
Application.ScreenUpdating = False
Dim strFile As String, StrDate As String, StrDir As String
Dim wb As Workbook
StrDir = MyPathFolder

StrDate = Format(Now, "dd-mm-yy")

strFile = Dir(StrDir & "*.csv")
Dim BA_Name As String
Do While Len(strFile) > 0
    If Right(strFile, 4) = ".csv" Then
        Set wb = Workbooks.Open(FileName:=StrDir & strFile, local:=True)
        Application.DisplayAlerts = False
        wb.SaveAs Replace(wb.FullName, ".csv", ".xls"), FileFormat:=xlExcel8
        Application.DisplayAlerts = True
        'remove data delegated to other advisors
        'BA_Name = wb.Name
        'remove_Confidential_Data (BA_Name)
        'close and save
        wb.Close True
        Set wb = Nothing
        strFile = Dir
    End If
Loop
Application.ScreenUpdating = True
End Sub


Sub Clear_All_CSV_files_In_Folder()
Application.ScreenUpdating = False
'Delete all files and subfolders
'Be sure that no file is open in the folder
    Dim FSO As Object
    Dim myPath As String

    Set FSO = CreateObject("scripting.filesystemobject")

    myPath = ThisWorkbook.Path

    If Right(myPath, 1) = "\" Then
        myPath = Left(myPath, Len(myPath) - 1)
    End If

    If FSO.FolderExists(myPath) = False Then
        'MsgBox MyPath & " doesn't exist"
        Exit Sub
    End If

    On Error Resume Next
    'Delete files
    FSO.deletefile myPath & "\*.csv", True
    'Delete subfolders
    'FSO.deletefolder myPath & "\*.*", True
   ' On Error GoTo 0
Application.ScreenUpdating = True
End Sub
Sub ClearFolder(foldername As String)
    'Delete all files and subfolders
'Be sure that no file is open in the folder
    Dim FSO As Object
    Dim myPath As String

    Set FSO = CreateObject("scripting.filesystemobject")

    myPath = foldername

    If Right(myPath, 1) = "\" Then
        myPath = Left(myPath, Len(myPath) - 1)
    End If

    If FSO.FolderExists(myPath) = False Then
        'MsgBox MyPath & " doesn't exist"
        Exit Sub
    End If

    On Error Resume Next
    'Delete files
    FSO.deletefile myPath & "\*.csv", True
    FSO.deletefile myPath & "\*.xls", True
    'Delete subfolders
    'FSO.deletefolder MyPath & "\*.*", True
    On Error GoTo 0
End Sub

Sub DeliverInputDataToChildSheets()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Locate the managers inputs and send to corresponding sheets
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets("Sheet1").Activate
Dim LastColumn As Integer: LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
Dim NextEmpty As Long: NextEmpty = Cells(Rows.Count, 20).End(xlUp).Row          'Business Name located in column T
Dim NextBA As Range: Set NextBA = Cells(NextEmpty, 20).Offset(1, -13)           'Bus Adviser located in column G
BusAdviser = NextBA.Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'set variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Target As Range: Set Target = NextBA.Offset(0, -6)
Target.Select
MyDate = Target.Offset(0, 1).Value
Jobnumber = Target.Offset(0, 2).Value
Topic = Target.Offset(0, 3).Value
LeadSource = Target.Offset(0, 4).Value
STATUS = Target.Offset(0, 5).Value
Program = Target.Offset(0, 6).Value
BusAdviser = Target.Offset(0, 7).Value
ClientID = Target.Offset(0, 8).Value
FirstName = Target.Offset(0, 9).Value
Surname = Target.Offset(0, 10).Value
Telephone = Target.Offset(0, 11).Value
Email = Target.Offset(0, 12).Value
Address = Target.Offset(0, 13).Value
Suburb = Target.Offset(0, 14).Value
State = Target.Offset(0, 15).Value
Postcode = Target.Offset(0, 16).Value
BusinessDuration = Target.Offset(0, 17).Value
ANZICCODE = Target.Offset(0, 18).Value
AboriginalBusiness = Target.Offset(0, 19).Value
BusinessName = Target.Offset(0, 20).Value
ABN = Target.Offset(0, 21).Value
CURRENTNumberofEmployees = Target.Offset(0, 22).Value
WomenInBusiness = Target.Offset(0, 23).Value
IndigenousInBusiness = Target.Offset(0, 24).Value
FamilyInBusiness = Target.Offset(0, 25).Value
Homebasedbusiness = Target.Offset(0, 26).Value
Fundingavenuesandfinancialanalysis = Target.Offset(0, 27).Value
Buildingyourbusiness = Target.Offset(0, 28).Value
Makingthemostofyourtalentandteam = Target.Offset(0, 29).Value
Managementcapabilities = Target.Offset(0, 30).Value
Digitalengagementimplementation = Target.Offset(0, 31).Value
Tourismready = Target.Offset(0, 32).Value
LegalNameofBusiness = Target.Offset(0, 33).Value
ABNforTHISbusiness = Target.Offset(0, 34).Value
SmallBusinessIntenderhasNOABN = Target.Offset(0, 35).Value
Consenttobesurveyed = Target.Offset(0, 36).Value
firstentrytoproject = Target.Offset(0, 37).Value
businessdiagnosticcompletedthisquarter = Target.Offset(0, 38).Value
UniqueBusiness = Target.Offset(0, 39).Value
SUST = Target.Offset(0, 40).Value
JobStartDate = Target.Offset(0, 41).Value
JobEndDate = Target.Offset(0, 42).Value
AgginreasedTO = Target.Offset(0, 43).Value
valuenewpositions = Target.Offset(0, 44).Value
Newpositions = Target.Offset(0, 45).Value
valueCapInv = Target.Offset(0, 46).Value
HOURSPROVIDED = Target.Offset(0, 47).Value
FollowUp = Target.Offset(0, 48).Value
Referredforwards = Target.Offset(0, 49).Value
InvDate = Target.Offset(0, 50).Value
Inv = Target.Offset(0, 51).Value
PaidDate = Target.Offset(0, 52).Value
Amount = Target.Offset(0, 53).Value
Notes = Target.Offset(0, 54).Value

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'look up BA child sheet to move data
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Workbooks.Open (ChildSheetLocation & NextBA & ".xls")

Cells(NextEmptyRow, 1).Value = MyDate
Cells(NextEmptyRow, 2).Value = Jobnumber
Cells(NextEmptyRow, 3).Value = Topic
Cells(NextEmptyRow, 4).Value = LeadSource
Cells(NextEmptyRow, 5).Value = STATUS
Cells(NextEmptyRow, 6).Value = Program
Cells(NextEmptyRow, 7).Value = BusAdviser
Cells(NextEmptyRow, 8).Value = ClientID
Cells(NextEmptyRow, 9).Value = FirstName
Cells(NextEmptyRow, 10).Value = Surname
Cells(NextEmptyRow, 11).Value = Telephone
Cells(NextEmptyRow, 12).Value = Email
Cells(NextEmptyRow, 13).Value = Address
Cells(NextEmptyRow, 14).Value = Suburb
Cells(NextEmptyRow, 15).Value = State
Cells(NextEmptyRow, 16).Value = Postcode
Cells(NextEmptyRow, 17).Value = BusinessDuration
Cells(NextEmptyRow, 18).Value = ANZICCODE
Cells(NextEmptyRow, 19).Value = AboriginalBusiness
Cells(NextEmptyRow, 20).Value = BusinessName
Cells(NextEmptyRow, 21).Value = ABN
Cells(NextEmptyRow, 22).Value = CURRENTNumberofEmployees
Cells(NextEmptyRow, 23).Value = WomenInBusiness
Cells(NextEmptyRow, 24).Value = IndigenousInBusiness
Cells(NextEmptyRow, 25).Value = FamilyInBusiness
Cells(NextEmptyRow, 26).Value = Homebasedbusiness
Cells(NextEmptyRow, 27).Value = Fundingavenuesandfinancialanalysis
Cells(NextEmptyRow, 28).Value = Buildingyourbusiness
Cells(NextEmptyRow, 29).Value = Makingthemostofyourtalentandteam
Cells(NextEmptyRow, 30).Value = Managementcapabilities
Cells(NextEmptyRow, 31).Value = Digitalengagementimplementation
Cells(NextEmptyRow, 32).Value = Tourismready
Cells(NextEmptyRow, 33).Value = LegalNameofBusiness
Cells(NextEmptyRow, 34).Value = ABNforTHISbusiness
Cells(NextEmptyRow, 35).Value = SmallBusinessIntenderhasNOABN
Cells(NextEmptyRow, 36).Value = Consenttobesurveyed
Cells(NextEmptyRow, 37).Value = firstentrytoproject
Cells(NextEmptyRow, 38).Value = businessdiagnosticcompletedthisquarter
Cells(NextEmptyRow, 39).Value = UniqueBusiness
Cells(NextEmptyRow, 40).Value = SUST
Cells(NextEmptyRow, 41).Value = JobStartDate
Cells(NextEmptyRow, 42).Value = JobEndDate
Cells(NextEmptyRow, 43).Value = AgginreasedTO
Cells(NextEmptyRow, 44).Value = valuenewpositions
Cells(NextEmptyRow, 45).Value = Newpositions
Cells(NextEmptyRow, 46).Value = valueCapInv
Cells(NextEmptyRow, 47).Value = HOURSPROVIDED
Cells(NextEmptyRow, 48).Value = FollowUp
Cells(NextEmptyRow, 49).Value = Referredforwards
Cells(NextEmptyRow, 50).Value = InvDate
Cells(NextEmptyRow, 51).Value = Inv
Cells(NextEmptyRow, 52).Value = PaidDate
Cells(NextEmptyRow, 53).Value = Amount
Cells(NextEmptyRow, 54).Value = Notes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'close and save changes to workbook
ActiveWorkbook.Close True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
Sub ObtainChildData()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Loop BA sheets to retreive data
'input data into Master
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Save New inputs and close master, on close update reporting sheets
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

