Attribute VB_Name = "templateCorrected_3"
Sub UpdatedRefreshTemplateWB()

'Name Workbooks to refer to
'set values from master workbook to transfer data
Dim wb As Workbook: Set wb = ThisWorkbook
Sheets("Sheet3").Activate
'Dim LastColumn As Integer: LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Len(Range("B4")) > 2 Then
    TemplatePath = Sheets("Sheet3").Cells(4, 2).Value
Else
    TemplatePath = InputBox("Please Insert Template File Path...", "Template Sheet Location", ThisWorkbook.path)
    'Sheets("Sheet3").Cells(1, 2).Value = TemplatePath
End If
If TemplatePath = "" Then
    MsgBox "Insert Path to Template Files"
    Exit Sub
End If
If TemplatePath = "" Then
    MsgBox "Insert Template Path": Exit Sub
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DECIDE IF MAC OR PC, USED WHILE TESTING
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Right(TemplatePath, 1) <> "\" Then
    If Right(TemplatePath, 1) <> ":" Then
    Cells(3, 2).Clear
        If IsMac = True Then
            Cells(4, 2).Clear
            If Right(TemplatePath, 1) <> ":" Then TemplatePath = TemplatePath & ":"
            Cells(3, 2).Value = "Mac"

        Else

            If IsMac = False Then
                Cells(4, 2).Clear
                If Right(TemplatePath, 1) <> "\" Then TemplatePath = TemplatePath & "\"
                Cells(3, 2).Value = "PC"
            End If
        End If
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
Dim TemplateWB As Workbook: Set TemplateWB = Workbooks.Open(TemplatePath & SBDC)


wb.Activate
Sheets("Sheet1").Activate
Application.ScreenUpdating = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim i As Integer
Dim LR As Integer: LR = Sheets("Sheet1").Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To LR

'4/19/17 error catching loop to only move unique business names to template sheet
UniqueOnly:
Sheets("Sheet1").Activate
Cells(i, "AN").Select
Selection.Copy
Selection.PasteSpecial xlPasteValues
'4/30/2017, Upper case added so 'new' and 'New' in column R on master would produce same results
If UCase(ActiveCell.Value) = "N" Then
    i = i + 1
    GoTo UniqueOnly
End If
'clear clipboard
Application.CutCopyMode = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Removed if statement 4/19/17
'4/10/17 added if statement to only transport unique values to data tab in SDBC report

        TemplateWB.Activate
        Sheets("Data").Select
        Dim LRdata As Integer: LRdata = Sheets("Data").Range("A9000").End(xlUp).Row
        wb.Activate
        Sheets("Sheet1").Activate
        'check to see if this is correct template
        If Cells(i, "F").Value <> "Business Local" Then GoTo NextRow
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Locate the inputs from the master sheet and send to corresponding values in template
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''***this causes all the flashing*****''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Activate Master Sheet find value then Activate SDBC Template and Insert Value
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        Title = Cells(i, 9).Value ': Debug.Print Title
        TemplateWB.Activate: Sheets("Data").Select: Range("A" & LRdata).Offset(1, 0).Value = Title
        wb.Activate: Sheets("Sheet1").Select
        
        FirstName = Cells(i, 10).Value ': Debug.Print FirstName
        TemplateWB.Activate: Sheets("Data").Select: Range("B9000").End(xlUp).Offset(1, 0).Value = FirstName
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
        
        
        ''''''CHANGED ABN TO V IN MASTER TO GO TO K IN SBDCReport
        '3/28/2017
        ABN = Cells(i, "V").Value '':debug.print ABN
        TemplateWB.Activate: Sheets("Data").Select: Range("K" & LRdata).Offset(1, 0).Value = ABN
        wb.Activate: Sheets("Sheet1").Select
        ''''''CHANGED Indig. Bus. TO T IN MASTER TO GO TO I IN SBDCReport
        '3/28/2017
        IndigenousInBusiness = Cells(i, "T").Value ':debug.print IndigenousInBusiness
        TemplateWB.Activate: Sheets("Data").Select: Range("I" & LRdata).Offset(1, 0).Value = IndigenousInBusiness
        
        'Removed autosave to prompt for save event after importing data
        'TemplateWB.Save
        
        wb.Activate
        Sheets("Sheet1").Activate
NextRow:
Next i
Application.ScreenUpdating = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PERFORMANCE TAB SDBC SHEET
wb.Activate
SDBCperformanceTab
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'FORMAT PHONE NUMBERS'
'Commented out, Not allowed to change telephone formatting in this template
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'4/19/17 Below FORMATS THE PHONE NUMBERS FOR THE SDBC TEMPLATE
'Sheets("Data").Select
'Dim TLR As Integer, D As Integer
'TLR = Cells(Rows.Count, "D").End(xlUp).Row

'For D = 3 To TLR
'    Cells(D, "D").Activate
'    If Left(LTrim(ActiveCell.Value), 1) <> 0 Then ActiveCell.Value = " " & 0 & ActiveCell.Value
'    ActiveCell = Replace(Replace(Replace(ActiveCell.Value, " ", ""), "-", ""), ".", "")
'    ActiveCell.NumberFormat = "0#""-""####""-""####"
'Next
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Prompt to save template
TemplateWB.Close 'True

wb.Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = True
If Application.EnableEvents = False Then Application.EnableEvents = True

'=============================================='
ASBAR 'CALL MACRO FOR ASBA TEMPLATE REPORT
'=============================================='

Application.ScreenUpdating = True
Application.DisplayAlerts = True


Exit Sub
TempUnfound:
If Err.Number = 1004 Then
    MsgBox "Could not locate SBDC Report Sheet, Please check the File Name and Path" & vbNewLine & vbNewLine & "SBDC Sheet name can be placed in Cell B5 on Sheet3", vbOKOnly, "Unable to Locate Template"
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
Application.EnableEvents = True
Dim wb As Workbook: Set wb = ThisWorkbook
Dim ASBAR As String
Sheets("Sheet3").Activate
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CHECK FOR TEMPLATE PATH, IF IT EXISTS CONTINUE, ELSE PROMPT FOR PATH
Dim TemplatePath As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Len(Range("B4")) > 2 Then
    TemplatePath = Sheets("Sheet3").Range("B4").Value
Else
    TemplatePath = InputBox("Please Insert Folder Path to Child Sheets", "BA Folder Location", ThisWorkbook.path)
    'Sheets("Sheet3").Cells(1, 2).Value = TemplatePath
End If
If TemplatePath = "" Then
    MsgBox "Insert Path to Import BA sheets"
    Exit Sub
End If
If TemplatePath = "" Then
    MsgBox "Insert Template Path": Exit Sub
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DECIDE IF MAC OR PC, USED WHILE TESTING
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Right(TemplatePath, 1) <> "\" Then
    If Right(TemplatePath, 1) <> ":" Then
    Cells(3, 2).Clear
        If IsMac = True Then
            Cells(4, 2).Clear
            If Right(TemplatePath, 1) <> ":" Then TemplatePath = TemplatePath & ":"
            Cells(3, 2).Value = "Mac"

        Else

            If IsMac = False Then
                Cells(4, 2).Clear
                If Right(TemplatePath, 1) <> "\" Then TemplatePath = TemplatePath & "\"
                Cells(3, 2).Value = "PC"
            End If
        End If
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'insert template path into sheet3
Sheets("Sheet3").Activate: Range("B4").Value = TemplatePath
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Locate Template Sheet ASBA from cell 'b6' or prompt for the name
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

If Len(Range("B6")) > 2 Then
    ASBAR = Range("B6").Value
Else
    ASBAR = InputBox("Enter ASBAS Template workbook name", "ASBAS", "ASBASReport.xlsx")
    Range("B6") = ASBAR
End If
If ASBAR = "" Then
    MsgBox "Insert ASBAS template sheet name...", vbOKOnly, "Please try again"
    Exit Sub
End If





On Error GoTo TempUnfound
Dim AB As Workbook: Set AB = Workbooks.Open(TemplatePath & ASBAR)
Dim ws As Worksheet: Set ws = Sheets("NATI client data")
ThisWorkbook.Activate
wb.Activate
Sheets("Sheet1").Activate


Dim LR As Integer: LR = Cells(Rows.Count, 1).End(xlUp).Row
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
Dim Prop_ind As String





Application.EnableEvents = False
For i = 2 To LR
Application.ScreenUpdating = False
AB.Activate
ws.Activate

'If the business name in column A is missing but data is entered, bounce to next row down
LRow = Cells(Rows.Count, 1).End(xlUp).Row
If Cells(LRow, 1).Offset(1, 0).Value = "" And Cells(LRow, 1).Offset(1, 1).Value <> "" Then LRow = LRow + 1

wb.Activate: Sheets("Sheet1").Activate

'Check to see if this is the correct template to use
If Cells(i, "F").Value <> "ASBAS NATI" Then GoTo asbaNextRow



Leg = Cells(i, 34).Value: AB.Activate: ws.Activate
Cells(LRow, 1).Offset(1, 0).Value = Leg

wb.Activate: Sheets("Sheet1").Activate

ABN = Cells(i, "AI").Value: AB.Activate: ws.Activate
Cells(LRow, 2).Offset(1, 0).Value = ABN

wb.Activate: Sheets("Sheet1").Activate

Sma = Cells(i, 36).Value: AB.Activate: ws.Activate
Cells(LRow, 3).Offset(1, 0).Value = Sma

wb.Activate: Sheets("Sheet1").Activate

Con = Cells(i, 37).Value: AB.Activate: ws.Activate
Cells(LRow, 4).Offset(1, 0).Value = Con

wb.Activate: Sheets("Sheet1").Activate

firstEntry = Cells(i, "AL").Value: AB.Activate: ws.Activate
Cells(LRow, 5).Offset(1, 0).Value = firstEntry



wb.Activate: Sheets("Sheet1").Activate

busDIAG = Cells(i, "AM").Value: AB.Activate: ws.Activate
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

FirstName = Cells(i, 10).Value:

AB.Activate: ws.Activate
Cells(LRow, 12).Offset(1, 0).Value = FirstName


wb.Activate
Sheets("Sheet1").Activate

Surname = Cells(i, "K").Value: AB.Activate: ws.Activate
Cells(LRow, "M").Offset(1, 0).Value = Surname


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
'''''''''''''''''''''''''''''''''''''
'CORRECTED
pop_ind = Cells(i, "S").Value: AB.Activate: ws.Activate
Cells(LRow, 16).Offset(1, 0).Value = pop_ind
'''''''''''''''''''''''''''''''''''''
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

wb.Activate
Sheets("Sheet1").Activate

asbaNextRow:

Next i

Application.ScreenUpdating = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'close and save changes to workbook
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
AB.Activate
'enable save events
AB.Close 'True
wb.Activate
If Application.EnableEvents = False Then Application.EnableEvents = True
Exit Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ERROR CATCHER
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
TempUnfound:
'Alert if wrong Template name was put in otherwise display alert
If Err.Number = 1004 Then
    MsgBox "Could not locate ABSA Report Sheet, Please check the File Name and Path" & vbNewLine & vbNewLine & "ABSA Sheet name can be placed in Cell B6 on Sheet3", vbOKOnly, "Unable to Locate Template"
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
Sub SDBCperformanceTab()
Application.DisplayAlerts = False
On Error GoTo NoBusNameSht
Sheets("Business Names").Delete
NoBusNameSht:
Application.DisplayAlerts = True

Sheets.Add: ActiveSheet.Name = "Business Names"
Cells(1, 1).Value = "Aboriginal": Cells(1, 2).Value = "Non Aboriginal": Cells(1, 3).Value = "Referrals"

Dim BN As Worksheet: Set BN = Sheets("Business Names")
Dim wb As Workbook: Set wb = ThisWorkbook

Dim NonAB As Integer, AB As Integer, LR As Integer, i As Integer
Dim ABMoney As Long, abVALnewpositions As Long, Num_abPos As Integer, nonABMoney As Long
Dim NonABVALnewPositions As Long, Num_NONabPos As Integer, BusRef As Integer, NonBusRef As Integer
Dim abCapInvested As Long, NONabCapInvested As Long, abHours As Integer, NonABhours As Integer
Dim Num_ABactionPlans As Integer, Num_NonABactionPlans As Long
Dim ABfollowUps As Integer, NonABFollowups As Integer, ServBus As Integer, NonServBus As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Changes 4/5/2017 via freelancer Sandeep
Dim abNewBus As Integer, NONabNewBus As Integer, abYearNEWbus As Integer, NONabYearNEWbus As Integer
Dim abPlan As Integer, NoNabPlan As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

AB = 0: NonAB = 0: ABMoney = 0: nonABMoney = o
With wb
Sheet1.Activate

LR = Cells(Rows.Count, "T").End(xlUp).Row
'Count the non aboriginal and aboriginal businesses
For i = 2 To LR
Sheets("Sheet1").Select
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'4/19/2017 identify refferred forwards, count non aboriginal reffered forwards in c15 of performance sheet
'otherwise count business name towards non aboriginal new business startups (b7 on performance sheet)
        If UCase(Trim(Cells(i, "AX").Value)) <> "NO" Then
            Cells(i, "U").Copy
            If UCase(Trim(Cells(i, "T").Value)) = "YES" Then
                BN.Select
                Range("C9000").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
            End If
            Application.CutCopyMode = False
        End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Sheet1").Select

    If UCase(Trim(Cells(i, "T").Value)) = "YES" Then
        AB = AB + 1
        ABMoney = ABMoney + Cells(i, "AR").Value
        abVALnewpositions = abVALnewpositions + Cells(i, "AS").Value
        Num_abPos = Num_abPos + Cells(i, "AT").Value
        abCapInvested = abCapInvested + Cells(i, "AU").Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Changes 4/5/2017 via freelancer Sandeep
        If UCase(Trim(Cells(i, "R").Value)) = "NEW" Then
            Cells(i, "R").Offset(0, 3).Copy
            BN.Select
            Range("A9000").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
        End If
        Sheets("Sheet1").Select
        If UCase(Trim(Cells(i, "AO"))) = "Y" Then abYearNEWbus = abYearNEWbus + 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        abHours = abHours + Cells(i, "AV").Value
        ABfollowUps = ABfollowUps + Cells(i, "AW").Value

        Sheets("Sheet1").Select
        If UCase(Trim(Cells(i, "AN").Value)) = "Y" Then ServBus = ServBus + 1
        If Len(Cells(i, "A")) > 0 Then abPlan = abPlan + 1
    Else
        If UCase(Trim(Cells(i, "T").Value)) = "NO" Then
            NonAB = NonAB + 1
            nonABMoney = nonABMoney + Cells(i, "AR").Value
            NonABVALnewPositions = NonABVALnewPositions + Cells(i, "AS").Value
            Num_NONabPos = Num_NONabPos + Cells(i, "AT").Value
            NONabCapInvested = NONabCapInvested + Cells(i, "AU").Value
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Changes 4/5/2017 via freelancer Sandeep
            If UCase(Trim(Cells(i, "R"))) = "NEW" Then
                Cells(i, "R").Offset(0, 3).Copy
                BN.Select
                Range("B9000").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
            End If
            Sheets("Sheet1").Select
            If Trim(Cells(i, "AO")) = "Y" Then NONabYearNEWbus = NONabYearNEWbus + 1
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            NonABhours = NonABhours + Cells(i, "AV").Value
            NonABFollowups = NonABFollowups + Cells(i, "AW").Value
            If Trim(Cells(i, "AN").Value) = "Y" Then NonServBus = NonServBus + 1
            If Len(Cells(i, "A")) > 0 Then NoNabPlan = NoNabPlan + 1
        End If
    End If
Application.CutCopyMode = False
Next i
End With

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'If template Workbook is open then activate it, otherwise open the workbook
'Get the default path to template and set variables
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets("Sheet3").Activate
Dim SDBCwb As Workbook
Dim MyName As String
MyName = Range("B5").Value
If WorkbookIsOpen(Range("B5").Value) = False Then
    Set SDBCwb = Workbooks.Open(Range("B4").Value & Range("B5").Value)
Else
    Set SDBCwb = Workbooks(MyName)
End If
SDBCwb.Activate
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'4/19/17 find correct sheet name dynamically as there are multiple spaces after performance that may get corrected
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Psh As Worksheet
For Each Psh In ActiveWorkbook.Worksheets
    If Left(Trim(Psh.Name), 11) = "Performance" Then Psh.Activate
Next Psh
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Range("B2").Value = nonABMoney
Range("C2").Value = ABMoney

Range("B3").Value = NonABVALnewPositions
Range("C3").Value = abVALnewpositions

Range("B4").Value = Num_NONabPos
Range("C4").Value = Num_abPos

Range("B5").Value = NONabCapInvested
Range("C5").Value = abCapInvested

Range("B8").Value = NONabYearNEWbus
Range("C8").Value = abYearNEWbus

Range("B13").Value = NoNabPlan
Range("C13").Value = abPlan

Range("B14").Value = NonABFollowups
Range("C14").Value = ABfollowUps

Range("B11").Value = NonServBus
Range("C11").Value = ServBus

Range("B12").Value = NonABhours
Range("C12").Value = abHours

'''''''''''''''''''''''''''''''''''''''''''''''
'4/19/17 updated
'Count Unique Business Names for Predetermined groups
'''''''''''''''''''''''''''''''''''''''''''''''
wb.Activate
BN.Activate
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
CountUniquesBusinessNames 'call sub to get count of unique business names for new business
'(aboriginal and non aboriginal) and business referrals
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
BN.Activate
abNewBus = Range("A900").End(xlUp).Value
NONabNewBus = Range("B900").End(xlUp).Value

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Changed 4/10/17 to mitigate error if all values for AX would equal "NO"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If Trim(Range("C2").Value) <> "" Then
    BusRef = Range("C900").End(xlUp).Value
Else
    BusRef = 0
End If


SDBCwb.Activate

Dim Performance As Worksheet
For Each Performance In ActiveWorkbook.Worksheets
    If Left(Trim(Performance.Name), 11) = "Performance" Then Performance.Activate
Next Performance

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'4/19/17 confirmed the macro is accurately identifying values
Range("B7").Value = NONabNewBus
Range("C7").Value = abNewBus
Range("C15").Value = BusRef
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'remove business names sheet after values have been transferred
wb.Activate
Application.DisplayAlerts = False
Sheets("Business Names").Delete
Application.DisplayAlerts = True
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SDBCwb.Activate
End Sub

Private Function WorkbookIsOpen(WBname) As Boolean
'returns true if workbook is open
Dim x As Workbook
On Error Resume Next
Set x = Workbooks(WBname)
If Err.Number = 0 Then
    WorkbookIsOpen = True
    Workbooks(WBname).Activate
Else
    WorkbookIsOpen = False
End If
End Function
Sub CountUniquesBusinessNames()

Sheets("Business Names").Activate
Dim ALR As Integer, BLR As Integer, CLR As Integer
ALR = Cells(Rows.Count, 1).End(xlUp).Row
BLR = Cells(Rows.Count, 2).End(xlUp).Row
CLR = Cells(Rows.Count, 3).End(xlUp).Row

If Len(Range("A2")) > 0 Then Range("A900").End(xlUp).Offset(1, 0).Value = CountUniqueVisible(Range("A2:A" & ALR))

If Len(Range("B2")) > 0 Then Range("B900").End(xlUp).Offset(1, 0).Value = CountUniqueVisible(Range("B2:B" & BLR))

If Len(Range("C2")) > 0 Then Range("C900").End(xlUp).Offset(1, 0).Value = CountUniqueVisible(Range("C2:C" & CLR))
End Sub
Function CountUniqueVisible(Target As Range) As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Function used by Allen to count unique names with correct attributes on business names sheet
'this sheet is a temp sheet which is deleted upon transferring calculated numbers
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Rng As Range, c As Range
Dim dic As Object
Dim y
Dim j As Long
Dim Sht As Worksheet
Dim strSheets As String
Set dic = CreateObject("Scripting.Dictionary")
Set Rng = Target.SpecialCells(xlCellTypeVisible)
j = 0
For Each c In Rng
    If Not dic.exists(c.Value) Then
        j = j + 1
        dic.Add c.Value, j
    End If
Next c

y = dic.keys
'Now we have a list of unique values. Next step is to return the count.
CountUniqueVisible = UBound(y) + 1
ExitHere:
Set dic = Nothing
Set Rng = Nothing
End Function
