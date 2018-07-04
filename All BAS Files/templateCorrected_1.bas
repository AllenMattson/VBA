Attribute VB_Name = "templateCorrected_1"
Sub RefreshTemplateWB()
Dim MyDate As String 'Date
Dim Jobnumber As String 'Single
Dim Topic As String 'String
Dim LeadSource As String 'String
Dim STATUS As String 'String
Dim Program As String 'String
Dim BusAdviser As String 'String
Dim ClientID As String 'String
Dim Title As String 'String
Dim FirstName As String 'String
Dim Surname As String 'String
Dim Telephone As String 'String
Dim Email As String 'String
Dim Address As String 'String
Dim Suburb As String 'String
Dim State As String 'String
Dim Postcode As String 'Integer
Dim BusinessDuration As String 'Integer
Dim ANZICCODE As String 'String
Dim AboriginalBusiness As String 'String
Dim BusinessName As String 'String
Dim ABN As String 'String
Dim CURRENTNumberofEmployees As String 'Integer
Dim WomenInBusiness As String 'String
Dim IndigenousInBusiness As String 'String
Dim FamilyInBusiness As String 'String
Dim Homebasedbusiness As String 'String
Dim Fundingavenuesandfinancialanalysis As String 'String
Dim Buildingyourbusiness As String 'String
Dim Makingthemostofyourtalentandteam As String 'String
Dim Managementcapabilities As String 'String
Dim Digitalengagementimplementation As String 'String
Dim Tourismready As String 'String
Dim LegalNameofBusiness As String 'String
Dim ABNforTHISbusiness As String 'String
Dim SmallBusinessIntenderhasNOABN As String 'String
Dim Consenttobesurveyed As String 'String
Dim firstentrytoproject As String 'String
Dim businessdiagnosticcompletedthisquarter As String 'String
Dim UniqueBusiness As String 'String
Dim SUST As String 'String
Dim JobStartDate As String 'Date
Dim JobEndDate As String 'Date
Dim AgginreasedTO As String 'Single
Dim valuenewpositions As String 'Single
Dim Newpositions As String 'Integer
Dim valueCapInv As String 'Single
Dim HOURSPROVIDED As String 'Integer
Dim FollowUp As String 'Integer
Dim Referredforwards As String 'String
Dim InvDate As String 'Date
Dim Inv As String 'Integer
Dim PaidDate As String 'Date
Dim Amount As String 'Single
Dim Notes As String 'String
'Name Workbooks to refer to
'set values from master workbook to transfer data
Dim wb As Workbook: Set wb = ThisWorkbook
Sheets("Sheet1").Activate
Dim LastColumn As Integer: LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'set variables
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim TemplatePath As String: TemplatePath = InputBox("Insert the file path to the template worksheets...", "Template file path", ThisWorkbook.path & "\")
If TemplatePath = "" Then
    MsgBox "Insert Template Path": Exit Sub
End If
    If Right(TemplatePath, 1) <> "\" Then
        If Right(TemplatePath, 1) <> ":" Then

            If IsMac = True Then
                Cells(4, 2).Clear
                If Right(TemplatePath, 1) <> ":" Then TemplatePath = TemplatePath & ":"
                Cells(4, 2).Value = TemplatePath
                Cells(3, 2).Clear
                Cells(3, 2).Value = "Mac"
            
            Else
                If IsMac = False Then
                    Cells(4, 2).Clear
                    If Right(TemplatePath, 1) <> "\" Then TemplatePath = TemplatePath & "\"
                    Cells(4, 2).Value = TemplatePath
                    Cells(3, 2).Clear
                    Cells(3, 2).Value = "PC"
                End If
            End If
        End If
    End If


Sheets("Sheet3").Activate: Range("B4").Value = TemplatePath


Dim SBDC As String: SBDC = InputBox("Enter SBDC Template workbook name", "SDBC", "SBDCReportingTemplateONE.xlsx")
Sheets("Sheet3").Activate
Range("B4") = TemplatePath
Range("B5") = SBDC


If SBDC = "" Then
    MsgBox "No SBDC Template name inserted": Exit Sub
End If
If Right(SBDC, 5) <> ".xlsx" Then SBDC = SBDC & ".xlsx"
On Error GoTo TempUnfound
Dim TemplateWB As Workbook: Set TemplateWB = Workbooks.Open(TemplatePath & SBDC)


wb.Activate
Sheets("Sheet1").Activate
Application.ScreenUpdating = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim i As Integer
Dim lr As Integer: lr = Sheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lr
TemplateWB.Activate
Sheets("Data").Select
Dim LRdata As Integer: LRdata = Sheets("Data").Cells(Rows.Count, 1).End(xlUp).Row
wb.Activate
Sheets("Sheet1").Activate
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Locate the managers inputs and send to corresponding sheets
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Activate SDBC Template and Insert Values
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Title = Cells(i, 9).Value ':debug.print Title
TemplateWB.Activate: Sheets("Data").Select: Range("A" & LRdata).Offset(1, 0).Value = Title
wb.Activate: Sheets("Sheet1").Select

FirstName = Cells(i, 10).Value ':debug.print FirstName
TemplateWB.Activate: Sheets("Data").Select: Range("B" & LRdata).Offset(1, 0).Value = FirstName
wb.Activate: Sheets("Sheet1").Select



Surname = Cells(i, 11).Value ':debug.print Surname
TemplateWB.Activate: Sheets("Data").Select: Range("C" & LRdata).Offset(1, 0).Value = Surname
wb.Activate: Sheets("Sheet1").Select



Telephone = Cells(i, 12).Value ':debug.print Telephone
TemplateWB.Activate: Sheets("Data").Select: Range("D" & LRdata).Offset(1, 0).Value = Telephone
wb.Activate: Sheets("Sheet1").Select



Email = Cells(i, 13).Value ':debug.print Email
TemplateWB.Activate: Sheets("Data").Select: Range("E" & LRdata).Offset(1, 0).Value = Email
wb.Activate: Sheets("Sheet1").Select


Address = Cells(i, 14).Value ':debug.print Address
Suburb = Cells(i, 15).Value ':debug.print Suburb
State = Cells(i, 16).Value ':debug.print State


Postcode = Cells(i, 17).Value ':debug.print Postcode
TemplateWB.Activate: Sheets("Data").Select: Range("F" & LRdata).Offset(1, 0).Value = Suburb & "/" & Postcode
wb.Activate: Sheets("Sheet1").Select



BusinessDuration = Cells(i, 18).Value ':debug.print BusinessDuration
TemplateWB.Activate: Sheets("Data").Select: Range("G" & LRdata).Offset(1, 0).Value = BusinessDuration
wb.Activate: Sheets("Sheet1").Select



ANZICCODE = Cells(i, 19).Value ':debug.print ANZICCODE
TemplateWB.Activate: Sheets("Data").Select: Range("H" & LRdata).Offset(1, 0).Value = ANZICCODE
wb.Activate: Sheets("Sheet1").Select


AboriginalBusiness = Cells(i, 20).Value ':debug.print AboriginalBusiness


BusinessName = Cells(i, 21).Value ':debug.print BusinessName
TemplateWB.Activate: Sheets("Data").Select: Range("J" & LRdata).Offset(1, 0).Value = BusinessName
wb.Activate: Sheets("Sheet1").Select



ABN = Cells(i, 22).Value '':debug.print ABN
TemplateWB.Activate: Sheets("Data").Select: Range("I" & LRdata).Offset(1, 0).Value = ABN
wb.Activate: Sheets("Sheet1").Select

IndigenousInBusiness = Cells(i, 25).Value ':debug.print IndigenousInBusiness
TemplateWB.Activate: Sheets("Data").Select: Range("K" & LRdata).Offset(1, 0).Value = IndigenousInBusiness

'TemplateWB.Save

wb.Activate
Sheets("Sheet1").Activate
Next i
Application.ScreenUpdating = True
'enable save event
TemplateWB.Close 'True

wb.Activate
Sheets("Sheet3").Activate

Range("B4").Value = TemplatePath

templateCorrected_1.ASBAR
Application.ScreenUpdating = True
Application.DisplayAlerts = True


Exit Sub
TempUnfound:
If Err.Number = 1004 Then
    MsgBox "Could not locate SBDC Report Sheet, Please check the File Name and Path", vbOKOnly, "Unable to Locate Template"
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
Else
    MsgBox "Error Alert!" & vbNewLine & "Error Number: " & Err.Number & vbNewLine & vbNewLine & "Description: " & vbNewLine & Err.Description, vbOKOnly, "Error in the macro!"
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub
End If

End Sub
Sub ASBAR()
Dim wb As Workbook: Set wb = ThisWorkbook
Dim TemplatePath As String
Sheets("Sheet3").Select
If Range("B4").Value <> "" Then
    TemplatePath = Range("B4").Value
Else
    TemplatePath = InputBox("Insert the file path to the template worksheets...", "Template file path", ThisWorkbook.path & "\")
        If TemplatePath = "" Then
            MsgBox "Insert Template Path": Exit Sub
        End If
    If Right(TemplatePath, 1) <> "\" Then
        If Right(TemplatePath, 1) <> ":" Then

            If IsMac = True Then
                Cells(4, 2).Clear
                If Right(TemplatePath, 1) <> ":" Then TemplatePath = TemplatePath & ":"
                Cells(4, 2).Value = TemplatePath
                Cells(3, 2).Clear
                Cells(3, 2).Value = "Mac"
            
            Else
                If IsMac = False Then
                    Cells(4, 2).Clear
                    If Right(TemplatePath, 1) <> "\" Then TemplatePath = TemplatePath & "\"
                    Cells(4, 2).Value = TemplatePath
                    Cells(3, 2).Clear
                    Cells(3, 2).Value = "PC"
                End If
            End If
        End If
    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' TemplatePath = InputBox("Insert the file path to the template worksheets...", "Template file path", ThisWorkbook.path & "\")
     '   If TemplatePath = "" Then
          '  MsgBox "Insert Template Path": Exit Sub
       ' End If
    'If Right(TemplatePath, 1) <> "\" Then TemplatePath = TemplatePath & "\": Range("B4").Value = TemplatePath

End If
Dim ASBAR As String: ASBAR = InputBox("Enter ASBAR Template workbook name", "ASBA", "ASBAReportingTemplate.xlsx")
If Right(ASBAR, 5) <> ".xlsx" Then ASBAR = ASBAR & ".xlsx"


Range("B6") = ASBAR


On Error GoTo TempUnfound
Dim AB As Workbook: Set AB = Workbooks.Open(TemplatePath & ASBAR)
Dim ws As Worksheet: Set ws = AB.Sheets("NATI client data")



wb.Activate
Sheets("Sheet1").Activate


Dim lr As Integer: lr = Cells(Rows.Count, 1).End(xlUp).Row
Dim LRow As Integer
''''ASBAR TEMPLATE
Dim Leg As String
Dim ABN As String
Dim Sma As String
Dim Con As String
Dim firstEntry As String
Dim busDIAG As String
Dim Address As String
Dim Suburb As String
Dim State As String
Dim Postcode As String
Dim Title As String
Dim FirstName As String
Dim Surname As String
Dim Telephone As String
Dim Email As String
Dim Industry As String
Dim CURRENTempl As String
Dim Wom As String
Dim Ind As String
Dim Fam As String
Dim Hom As String
Dim Fun As String
Dim Bui As String
Dim Mak As String
Dim Man As String
Dim Dig As String
Dim Tou As String





Application.ScreenUpdating = False
Application.EnableEvents = False
For i = 2 To lr
AB.Activate
ws.Activate
LRow = Cells(Rows.Count, 1).End(xlUp).Row
If Cells(LRow, 1).Offset(1, 0).Value = "" And Cells(LRow, 1).Offset(1, 1).Value <> "" Then LRow = LRow + 1
wb.Activate: Sheets("Sheet1").Activate
Leg = Cells(i, 34).Value: AB.Activate: ws.Activate
Cells(LRow, 1).Offset(1, 0).Value = Leg

wb.Activate: Sheets("Sheet1").Activate

ABN = Cells(i, 22).Value: AB.Activate: ws.Activate
Cells(LRow, 2).Offset(1, 0).Value = ABN

wb.Activate: Sheets("Sheet1").Activate

Sma = Cells(i, 36).Value: AB.Activate: ws.Activate
Cells(LRow, 3).Offset(1, 0).Value = Sma

wb.Activate: Sheets("Sheet1").Activate

Con = Cells(i, 37).Value: AB.Activate: ws.Activate
Cells(LRow, 4).Offset(1, 0).Value = Con

wb.Activate: Sheets("Sheet1").Activate

firstEntry = Cells(i, 10).Value: AB.Activate: ws.Activate
Cells(LRow, 5).Offset(1, 0).Value = firstEntry



wb.Activate: Sheets("Sheet1").Activate

busDIAG = Cells(i, 18).Value: AB.Activate: ws.Activate
Cells(LRow, 6).Offset(1, 0).Value = busDIAG


wb.Activate: Sheets("Sheet1").Activate

Postcode = Cells(i, 17).Value: AB.Activate: ws.Activate
Cells(LRow, 10).Offset(1, 0).Value = Postcode


wb.Activate: Sheets("Sheet1").Activate

Address = Cells(i, 14).Value: AB.Activate: ws.Activate
Cells(LRow, 7).Offset(1, 0).Value = Address



wb.Activate
Sheets("Sheet1").Activate

Suburb = Cells(i, 15).Value: AB.Activate: ws.Activate
Cells(LRow, 8).Offset(1, 0).Value = Suburb

wb.Activate
Sheets("Sheet1").Activate

State = Cells(i, 16).Value: AB.Activate: ws.Activate
Cells(LRow, 9).Offset(1, 0).Value = State


wb.Activate
Sheets("Sheet1").Activate

Title = Cells(i, 9).Value: AB.Activate: ws.Activate
Cells(LRow, 11).Offset(1, 0).Value = Title


wb.Activate
Sheets("Sheet1").Activate

FirstName = Cells(i, 10).Value: AB.Activate: ws.Activate
Cells(LRow, 12).Offset(1, 0).Value = FirstName


wb.Activate
Sheets("Sheet1").Activate

Sur = Cells(i, 11).Value: AB.Activate: ws.Activate
Cells(LRow, 13).Offset(1, 0).Value = Surname


wb.Activate
Sheets("Sheet1").Activate

Telephone = Cells(i, 12).Value: AB.Activate: ws.Activate
Cells(LRow, 14).Offset(1, 0).Value = Telephone


wb.Activate
Sheets("Sheet1").Activate

Email = Cells(i, 13).Value: AB.Activate: ws.Activate
Cells(LRow, 15).Offset(1, 0).Value = Email







wb.Activate
Sheets("Sheet1").Activate


CURRENTempl = Cells(i, 23).Value: AB.Activate: ws.Activate
Cells(LRow, 17).Offset(1, 0).Value = CURRENTempl

wb.Activate
Sheets("Sheet1").Activate

Wom = Cells(i, 24).Value: AB.Activate: ws.Activate
Cells(LRow, 18).Offset(1, 0).Value = Wom

wb.Activate
Sheets("Sheet1").Activate

Ind = Cells(i, 25).Value: AB.Activate: ws.Activate
Cells(LRow, 19).Offset(1, 0).Value = Ind

wb.Activate
Sheets("Sheet1").Activate

Fam = Cells(i, 26).Value: AB.Activate: ws.Activate
Cells(LRow, 20).Offset(1, 0).Value = Fam

wb.Activate
Sheets("Sheet1").Activate

Hom = Cells(i, 27).Value: AB.Activate: ws.Activate
Cells(LRow, 21).Offset(1, 0).Value = Hom

wb.Activate
Sheets("Sheet1").Activate

Fun = Cells(i, 28).Value: AB.Activate: ws.Activate
Cells(LRow, 22).Offset(1, 0).Value = Fun

wb.Activate
Sheets("Sheet1").Activate

Bui = Cells(i, 29).Value: AB.Activate: ws.Activate
Cells(LRow, 23).Offset(1, 0).Value = Bui

wb.Activate
Sheets("Sheet1").Activate

Mak = Cells(i, 30).Value: AB.Activate: ws.Activate
Cells(LRow, 24).Offset(1, 0).Value = Mak

wb.Activate
Sheets("Sheet1").Activate

Man = Cells(i, 31).Value: AB.Activate: ws.Activate
Cells(LRow, 25).Offset(1, 0).Value = Man

wb.Activate
Sheets("Sheet1").Activate

Dig = Cells(i, 32).Value: AB.Activate: ws.Activate
Cells(LRow, 26).Offset(1, 0).Value = Dig

wb.Activate
Sheets("Sheet1").Activate

Tou = Cells(i, 33).Value: AB.Activate: ws.Activate
Cells(LRow, 27).Offset(1, 0).Value = Tou

'AB.Save True

wb.Activate
Sheets("Sheet1").Activate


Next i




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'close and save changes to workbook
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

AB.Activate
'enable save events
AB.Close 'True
wb.Activate
Application.EnableEvents = True
Exit Sub
TempUnfound:
If Err.Number = 1004 Then
    MsgBox "Could not locate ABSA Report Sheet, Please check the File Name and Path", vbOKOnly, "Unable to Locate Template"
    Exit Sub
Else
    MsgBox "Error Alert!" & vbNewLine & "Error Number: " & Err.Number & vbNewLine & vbNewLine & "Description: " & vbNewLine & Err.Description, vbOKOnly, "Error in the macro!"
    Exit Sub
End If
End Sub
Function IsMac() As Boolean
#If Mac Then
    IsMac = True
#ElseIf Win32 Or Win64 Then
    IsMac = False
#End If
End Function
