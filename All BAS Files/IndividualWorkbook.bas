Attribute VB_Name = "IndividualWorkbook"
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
Sub DeliverInputDataToChildSheets()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Locate the managers inputs and send to corresponding sheets
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets("Sheet1").Activate
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
Workbooks.Open (NextBA & ".xls")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
Sub ObtainChildData()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Loop BA sheets and input data into Master
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




End Sub
