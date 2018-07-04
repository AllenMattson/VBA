Attribute VB_Name = "UpdateTemplates"
Sub RefreshTemplateWB()
Dim MyDate As Date
Dim Jobnumber As Single
Dim Topic As String
Dim LeadSource As String
Dim STATUS As String
Dim Program As String
Dim BusAdviser As String
Dim ClientID As String
Dim Title As String
Dim FirstName As String
Dim Surname As String
Dim Telephone As String
Dim Email As String
Dim Address As String
Dim Suburb As String
Dim State As String
Dim Postcode As Integer
Dim BusinessDuration As Integer
Dim ANZICCODE As String
Dim AboriginalBusiness As String
Dim BusinessName As String
Dim ABN As String
Dim CURRENTNumberofEmployees As Integer
Dim WomenInBusiness As String
Dim IndigenousInBusiness As String
Dim FamilyInBusiness As String
Dim Homebasedbusiness As String
Dim Fundingavenuesandfinancialanalysis As String
Dim Buildingyourbusiness As String
Dim Makingthemostofyourtalentandteam As String
Dim Managementcapabilities As String
Dim Digitalengagementimplementation As String
Dim Tourismready As String
Dim LegalNameofBusiness As String
Dim ABNforTHISbusiness As String
Dim SmallBusinessIntenderhasNOABN As String
Dim Consenttobesurveyed As String
Dim firstentrytoproject As String
Dim businessdiagnosticcompletedthisquarter As String
Dim UniqueBusiness As String
Dim SUST As String
Dim JobStartDate As Date
Dim JobEndDate As Date
Dim AgginreasedTO As Single
Dim valuenewpositions As Single
Dim Newpositions As Integer
Dim valueCapInv As Single
Dim HOURSPROVIDED As Integer
Dim FollowUp As Integer
Dim Referredforwards As String
Dim InvDate As Date
Dim Inv As Integer
Dim PaidDate As Date
Dim Amount As Single
Dim Notes As String
'Name Workbooks to refer to
'set values from master workbook to transfer data
Dim WB As Workbook: Set WB = ThisWorkbook
Sheets("Sheet1").Activate
Dim LastColumn As Integer: LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'set variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim TemplatePath As String: TemplatePath = InputBox("Insert the file path to the template worksheets...", "Template file path", ThisWorkbook.path)
Dim ASBAR As String: ASBAR = InputBox("Enter ASBAR Template workbook name", "ASBAR", "ASBARReportingTemplateONE.xlsx")
Dim SBDC As String: SBDC = InputBox("Enter ASBAR Template workbook name", "ASBAR", "SBDCReportingTemplateONE.xlsx")
If Right(ASBAR, 5) <> ".xlsx" Then ASBAR = ASBAR & ".xlsx"
If Right(SBDC, 5) <> ".xlsx" Then SBDC = sdbc & ".xlsx"
Dim errMsg As String: errMsg = "This folder does not containt a sheet"
On Error GoTo ErrornoTemplateFile
Dim ASBARwb As Workbook: Set ASBARwb = Workbooks.Open(ASBAR)
Dim TemplateWB As Workbook: Set TemplateWB = Workbooks.Open(SBDC)
ErrornoTemplateFile:
If Err.Number > 0 Or Err.Number < 0 Then
    MsgBox "Oh, hey.." & vbNewLine & "You Can't Locate Templates, Please Check the file path" & vbNewLine & " And Template Book names." & vbNewLine & "Unable to locate: " & vbNewLine & vbNewLine & ASBAR & vbNewLine & SBDC & vbNewLine & vbNewLine & "Checked in file path: " & TemplatePath, vbExclamation, "Missing Templates"
    Exit Sub
Else
NoErrorHere:
    'continue with code
End If

WB.Activate
Sheets("Sheet1").Activate
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim i As Integer
Dim LR As Integer: LR = Sheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To LR
WB.Activate
Sheets("Sheet1").Activate

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Locate the managers inputs and send to corresponding sheets
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 
Dim Target As Range: Set Target = Cells(1, 1)
Target.Select
MyDate = Target.Offset(i, 1).Value
Jobnumber = Target.Offset(i, 2).Value
Topic = Target.Offset(i, 3).Value
LeadSource = Target.Offset(i, 4).Value
STATUS = Target.Offset(i, 5).Value
Program = Target.Offset(i, 6).Value
BusAdviser = Target.Offset(i, 7).Value
ClientID = Target.Offset(i, 8).Value
Title = Target.Offset(i, 9).Value
FirstName = Target.Offset(i, 10).Value
Surname = Target.Offset(i, 11).Value
Telephone = Target.Offset(i, 12).Value
Email = Target.Offset(i, 13).Value
Address = Target.Offset(i, 14).Value
Suburb = Target.Offset(i, 15).Value
State = Target.Offset(i, 16).Value
Postcode = Target.Offset(i, 17).Value
BusinessDuration = Target.Offset(i, 18).Value
ANZICCODE = Target.Offset(i, 19).Value
AboriginalBusiness = Target.Offset(i, 20).Value
BusinessName = Target.Offset(i, 21).Value
ABN = Target.Offset(i, 22).Value
CURRENTNumberofEmployees = Target.Offset(i, 23).Value
WomenInBusiness = Target.Offset(i, 24).Value
IndigenousInBusiness = Target.Offset(i, 25).Value
FamilyInBusiness = Target.Offset(i, 26).Value
Homebasedbusiness = Target.Offset(i, 27).Value
Fundingavenuesandfinancialanalysis = Target.Offset(i, 28).Value
Buildingyourbusiness = Target.Offset(i, 29).Value
Makingthemostofyourtalentandteam = Target.Offset(i, 30).Value
Managementcapabilities = Target.Offset(i, 31).Value
Digitalengagementimplementation = Target.Offset(i, 32).Value
Tourismready = Target.Offset(i, 33).Value
LegalNameofBusiness = Target.Offset(i, 34).Value
ABNforTHISbusiness = Target.Offset(i, 35).Value
SmallBusinessIntenderhasNOABN = Target.Offset(i, 36).Value
Consenttobesurveyed = Target.Offset(i, 37).Value
firstentrytoproject = Target.Offset(i, 38).Value
businessdiagnosticcompletedthisquarter = Target.Offset(i, 39).Value
UniqueBusiness = Target.Offset(i, 40).Value
SUST = Target.Offset(i, 41).Value
JobStartDate = Target.Offset(i, 42).Value
JobEndDate = Target.Offset(i, 43).Value
AgginreasedTO = Target.Offset(i, 44).Value
valuenewpositions = Target.Offset(i, 45).Value
Newpositions = Target.Offset(i, 46).Value
valueCapInv = Target.Offset(i, 47).Value
HOURSPROVIDED = Target.Offset(i, 48).Value
FollowUp = Target.Offset(i, 49).Value
Referredforwards = Target.Offset(i, 50).Value
InvDate = Target.Offset(i, 51).Value
Inv = Target.Offset(i, 52).Value
PaidDate = Target.Offset(i, 53).Value
Amount = Target.Offset(i, 54).Value
Notes = Target.Offset(i, 55).Value


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Templates and Insert Values
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
TemplateWB.Activate
Sheets("Data").Activate
With TemplateWB
Dim LRdata As Integer: LRdata = Sheets("Data").Cells(Rows.Count, 1).End(xlUp).Row
Dim j As Long

Title = Range("A" & LRdata).Offset(1, 0).Value
FirstName = Range("B" & LRdata).Offset(1, 0).Value
Surname = Range("C" & LRdata).Offset(1, 0).Value
Telephone = Range("D" & LRdata).Offset(1, 0).Value
Email = Range("E" & LRdata).Offset(1, 0).Value
Postcode = Range("F" & LRdata).Offset(1, 0).Value
BusinessDuration = Range("G" & LRdata).Offset(1, 0).Value
ANZICCODE = Range("H" & LRdata).Offset(1, 0).Value
ABN = Range("I" & LRdata).Offset(1, 0).Value
BusinessName = Range("J" & LRdata).Offset(1, 0).Value
IndigenousInBusiness = Range("K" & LRdata).Offset(1, 0).Value
End With


''''ASBAR TEMPLATE
ASBARwb.Activate
Sheets("NATI client data").Activate
With ASBARwb
Dim LRRdata As Integer: LRRdata = Sheets("NATI client data").Cells(Rows.Count, 1).End(xlUp).Row
Dim LL As Long
MyDate = Range("A" & LRRdata).Offset(1, 0).Value
Jobnumber = Range("B" & LRRdata).Offset(1, 0).Value
Topic = Range("C" & LRRdata).Offset(1, 0).Value
LeadSource = Range("D" & LRRdata).Offset(1, 0).Value
STATUS = Range("E" & LRRdata).Offset(1, 0).Value
Program = Range("F" & LRRdata).Offset(1, 0).Value
BusAdviser = Range("G" & LRRdata).Offset(1, 0).Value
ClientID = Range("H" & LRRdata).Offset(1, 0).Value
Title = Range("I" & LRRdata).Offset(1, 0).Value
FirstName = Range("J" & LRRdata).Offset(1, 0).Value
Surname = Range("K" & LRRdata).Offset(1, 0).Value
Telephone = Range("L" & LRRdata).Offset(1, 0).Value
Email = Range("M" & LRRdata).Offset(1, 0).Value
Address = Range("N" & LRRdata).Offset(1, 0).Value
Suburb = Range("O" & LRRdata).Offset(1, 0).Value
State = Range("P" & LRRdata).Offset(1, 0).Value
Postcode = Range("Q" & LRRdata).Offset(1, 0).Value
BusinessDuration = Range("R" & LRRdata).Offset(1, 0).Value
ANZICCODE = Range("S" & LRRdata).Offset(1, 0).Value
AboriginalBusiness = Range("T" & LRRdata).Offset(1, 0).Value
BusinessName = Range("U" & LRRdata).Offset(1, 0).Value
ABN = Range("V" & LRRdata).Offset(1, 0).Value
CURRENTNumberofEmployees = Range("W" & LRRdata).Offset(1, 0).Value
WomenInBusiness = Range("X" & LRRdata).Offset(1, 0).Value
IndigenousInBusiness = Range("Y" & LRRdata).Offset(1, 0).Value
FamilyInBusiness = Range("Z" & LRRdata).Offset(1, 0).Value
Homebasedbusiness = Range("AA" & LRRdata).Offset(1, 0).Value
Fundingavenuesandfinancialanalysis = Range("" & LRRdata).Offset(1, 0).Value
Buildingyourbusiness = Range("" & LRRdata).Offset(1, 0).Value
Makingthemostofyourtalentandteam = Range("" & LRRdata).Offset(1, 0).Value
Managementcapabilities = Range("" & LRRdata).Offset(1, 0).Value
Digitalengagementimplementation = Range("" & LRRdata).Offset(1, 0).Value
Tourismready = Range("" & LRRdata).Offset(1, 0).Value
LegalNameofBusiness = Range("" & LRRdata).Offset(1, 0).Value
ABNforTHISbusiness = Range("" & LRRdata).Offset(1, 0).Value
SmallBusinessIntenderhasNOABN = Range("" & LRRdata).Offset(1, 0).Value
Consenttobesurveyed = Range("" & LRRdata).Offset(1, 0).Value
firstentrytoproject = Range("" & LRRdata).Offset(1, 0).Value
businessdiagnosticcompletedthisquarter = Range("" & LRRdata).Offset(1, 0).Value
UniqueBusiness = Range("" & LRRdata).Offset(1, 0).Value
SUST = Range("" & LRRdata).Offset(1, 0).Value
JobStartDate = Range("" & LRRdata).Offset(1, 0).Value
JobEndDate = Range("" & LRRdata).Offset(1, 0).Value
AgginreasedTO = Range("" & LRRdata).Offset(1, 0).Value
valuenewpositions = Range("" & LRRdata).Offset(1, 0).Value
Newpositions = Range("" & LRRdata).Offset(1, 0).Value
valueCapInv = Range("" & LRRdata).Offset(1, 0).Value
HOURSPROVIDED = Range("" & LRRdata).Offset(1, 0).Value
FollowUp = Range("" & LRRdata).Offset(1, 0).Value
Referredforwards = Range("" & LRRdata).Offset(1, 0).Value
InvDate = Range("" & LRRdata).Offset(1, 0).Value
Inv = Range("" & LRRdata).Offset(1, 0).Value
PaidDate = Range("" & LRRdata).Offset(1, 0).Value
Amount = Range("" & LRRdata).Offset(1, 0).Value
Notes = Range("" & LRRdata).Offset(1, 0).Value

End With

Next i

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'close and save changes to workbook
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
TemplateWB.Close True
ASBARwb.Close True
WB.Activate
Application.DisplayAlerts = True
End Sub
