Attribute VB_Name = "JEM_Main"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MODULE: JEM_MAIN
'AUTHOR: Allen Mattson
'DATE: 5/8/2018
'DESCRIPTION:
'   Use of commenting in code is kept to minimal, information on the code is provided at the top of the module, otherwise code is commented accordingly.
'   Variable names are to be descriptive of their purpose, if the macro is long, the variables are declared where they are used to avoid confusion to its purpose.
'   All toolbar interactions are stored in JemRibboN and reference macros located in the JEM_Main Module of the JEM2018 vba project.
'   Macros that are called in the JEM_Main module control printing, views and creating new journals.
'   When validate entries is called (via the ribbon) the module called is JEM_ValidateEntries.
'       **All macros regarding entries and calling database validation is checked here.
'   JEM_ValidateRecords is a module called in JEM_ValidateEntries and handles database validation and connections. Global variables are declared here.
'   btnBalanceAll is called when the 'BALANCE ALL' button is pressed by the user.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Private Ctot As Double 'credit total
Private Dtot As Double 'debit total
Private Btot As Currency
Sub ShowTemplate()
Sheets("newJE").Visible = xlSheetVisible
Sheets("newJE").Activate
End Sub
Sub HideTemplate()
Sheets("HBI").Activate
Sheets("newJE").Visible = xlSheetVeryHidden
End Sub
Sub CreateNewJournal()
'Confirm user wants to make new journal before proceding
Dim MSG As String, Mprompt As String, mTitle As String
MSG = "Do you want to open a new Journal workbook?"
Mprompt = "vbyesno+vbquestion+vbdefaultbutton2"
mTitle = "New Workbook"
Dim Answer As Integer
Answer = MsgBox(MSG, vbYesNo + vbQuestion + vbDefaultButton2, mTitle)
If Answer <> 6 Then Exit Sub '6 is returned if user selects yes

Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationAutomatic
Dim NewTab As String
Dim NewJE As Worksheet

Dim sh As Worksheet
For Each sh In ThisWorkbook.Worksheets
    If sh.Name = "newJE" Then
        sh.Visible = xlSheetVisible
        Set NewJE = sh
        NewJE.Copy
        Application.Dialogs(xlDialogSaveAs).Show
        ActiveSheet.Name = InputBox("Please enter a name for the new tab", "New Journal Created", "New Journal")
        NewJE.Visible = xlSheetVeryHidden
        GoTo FinishedNewJESheet
    End If
Next
FinishedNewJESheet:
'set the focus to new worksheet
'Place balance button into new worksheet
JEM_Main.MakeBalanceButton
ActiveSheet.Range(Columns(1), Columns(12)).Select
ActiveWindow.Zoom = True
Cells(1, 1).Select
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Cells(1, 1).Activate
End Sub
Sub PrintOut()

HideBlankRows

With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .DisplayAlerts = True
    .Dialogs(xlDialogPrint).Show
End With

UnhideRows

End Sub
Sub PrintPreview()

HideBlankRows

With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .DisplayAlerts = True
    .Dialogs(xlDialogPrintPreview).Show
End With

UnhideRows

End Sub
Sub SynchView(Optional View2 As Integer)
Application.ScreenUpdating = False
If View2 > 0 Then
    ActiveSheet.Range(Columns(1), Columns(View2)).Select
    ActiveWindow.Zoom = True
    Cells(1, 1).Select
Else
    ActiveSheet.Range(Columns(1), Columns("L")).Select
    ActiveWindow.Zoom = True
    Cells(1, 1).Select
End If
Application.ScreenUpdating = True
End Sub
Private Sub HideBlankRows()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False

Dim xRg As Range
Dim xCell As Range
Dim xAddress As String
Dim xUpdate As Boolean
Dim i As Long

On Error Resume Next

xAddress = Application.ActiveWindow.RangeSelection.Address
Set xRg = ActiveSheet.Range("A3:J1001")
Set xRg = Application.Intersect(xRg, ActiveSheet.UsedRange)

If xRg Is Nothing Then GoTo noxRG

For i = xRg.Rows.count To 1 Step -1
    xRg.Rows(i).EntireRow.Hidden = (Application.CountA(xRg.Rows(i)) = 0)
Next

noxRG:

Application.ScreenUpdating = True
Application.Calculation = xlAutomatic
Application.DisplayAlerts = True
End Sub
Private Sub UnhideRows()
Dim xRg As Range
Dim xCell As Range
Dim xAddress As String
Dim xUpdate As Boolean
On Error Resume Next
xAddress = Application.ActiveWindow.RangeSelection.Address
Set xRg = ActiveSheet.Range("A3:J" & Rows.count)
Set xRg = Application.Intersect(xRg, ActiveSheet.UsedRange)
If xRg Is Nothing Then Exit Sub
xUpdate = Application.ScreenUpdating
Application.ScreenUpdating = False
xRg.EntireRow.Hidden = False
Application.ScreenUpdating = xUpdate
End Sub
Private Sub MakeBalanceButton()

On Error GoTo MakeButton
Dim oSH As Shape
For Each oSH In ActiveSheet.Shapes
    If oSH.Name = "btnBalanceAll" Then oSH.Delete
Next oSH

MakeButton:
On Error GoTo 0
    Columns("I:I").Select
    ActiveSheet.buttons.Add(751.5, 0, 129, 27.75).Select
    Selection.OnAction = "JEM2018.xlam!btnBalanceAll"
    Selection.Characters.Text = "BALANCE ALL"
    Selection.Name = "btnBalanceAll"
    With Selection.Characters(Start:=1, Length:=11)
        With .Font
            .Name = "Calibri"
            .FontStyle = "Regular"
            .Bold = True
            .Size = 10
            .ColorIndex = 5
        End With
    End With
    Range("K1").Select
End Sub
Sub btnBalanceAll()
If Range("F1").Value <> "BALANCE" Then
    MsgBox "Select Journal Entry Worksheet"
    Exit Sub
End If
ActiveSheet.Range("K6:K1000").ClearContents
On Error GoTo NoValueFound
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Call sub to go group by group in table entries to format any
'   unbalanced groups
'Pass all values from credit range and debit range
'If false, change totals boxes backgrounds to blue
'Highlight the group whose balance <> 0
Call JEM_Main.LocateUnabalancedGroup
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Ctot = 0
Dtot = 0
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''This macro changes colors depending'''''''''''''''
'''''''''''''if the credits and debits balance''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Credits As Range, Debits As Range, TotalRange As Range
Dim Clr As Long
Dim Dlr As Long
ActiveSheet.Range("J6:J100000").SpecialCells(xlCellTypeConstants, 1).Select
Set Credits = Selection
ActiveSheet.Range("I6:I100000").SpecialCells(xlCellTypeConstants, 1).Select
Set Debits = Selection

Dim TotalsBoxes As Range
Set TotalsBoxes = Range("C1, E1, H1")
TotalsBoxes.ClearContents
Set TotalRange = Range(Credits, Debits)


Credits.NumberFormat = "$#,##0.00_);($#,##0.00)"
Debits.NumberFormat = "$#,##0.00_);($#,##0.00)"
'Put entry totals in their corresponding cells
'format the ranges
If BalanceCredDeb(Credits, Debits) = False Then
    With TotalsBoxes
        .Interior.Color = vbBlue
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 16
    End With
Else
    With TotalsBoxes
        .Interior.Color = vbWhite
        .Font.Color = vbBlack
        .Font.Bold = False
        .Font.Size = 12
        .Value = ""
    End With
End If
Btot = Ctot - Dtot
With ActiveSheet
    .Range("C1").Formula = "=Round(" & Dtot & ", 2)"
    .Range("E1").Formula = "=Round(" & Ctot & ", 2)"
    .Range("H1").Value = Btot
End With
Exit Sub

'Set everything to normal, alert user if no entry values found in debits or credits
NoValueFound:
If Err.Number = 1004 Then MsgBox "No Credit or Debit Found" & vbNewLine & vbNewLine & "Error: " & Err.Number & vbNewLine & Err.Description, vbInformation + vbOKOnly, "Credit or Debit Needed"
With ActiveSheet.Range("J6:J10000").Font
    .Color = vbBlack
    .Bold = False
End With
With Range("I6:I10000").Font
    .Color = vbBlack
    .Bold = False
End With
With TotalsBoxes
    .Interior.Color = vbWhite
    .Font.Color = vbBlack
    .Font.Bold = False
    .Font.Size = 12
    .Value = ""
End With
End Sub


Private Sub LocateUnabalancedGroup()
Application.ScreenUpdating = False
Dim Credits As Range, Debits As Range
Rows("1000:1000").EntireRow.Hidden = True
Dim iLR As Long, i As Long
Dim Dlr As Long, Clr As Long
Dim Alr As Long, k As Long
On Error GoTo ErrHandler
NextGroup:
Alr = Cells(Rows.count, 1).End(xlUp).row
If Cells(Alr, 1).Value = "Description" Then GoTo ErrHandler
Dlr = Cells(Rows.count, "I").End(xlUp).row
Clr = Cells(Rows.count, "J").End(xlUp).row
If Dlr > Clr Then Cells(Dlr, "I").Select
If Clr > Dlr Then Cells(Clr, "J").Select
If Dlr = Clr Then Cells(Clr, "J").Select

Set Debits = Range("I" & Alr & ":" & "I" & Selection.row)
Set Credits = Range("J" & Alr & ":" & "J" & Selection.row)

If BalanceCredDeb(Credits, Debits) = False Then
    With Range("I" & Alr & ":J" & Selection.row).Cells.Font
        .Color = vbBlue
        .Bold = True
    End With
Else
    With Range("I" & Alr & ":J" & Selection.row).Cells.Font
        .Color = vbBlack
        .Bold = False
    End With
End If
Debug.Print "Top: " & Alr & " bottom: " & Selection.row
For k = Selection.row To Alr Step -1
    Rows(k).EntireRow.Hidden = True
Next k

GoTo NextGroup

ErrHandler:
If Err.Number <> 0 Then MsgBox Err.Number & vbNewLine & vbNewLine & Err.Description
ActiveSheet.Rows.Hidden = False
Application.ScreenUpdating = True
End Sub
Private Function BalanceCredDeb(CredRNG As Range, DebRNG As Range) As Boolean
Ctot = 0
Dtot = 0
'Add numeric values in range and make sure they balance
BalanceCredDeb = True
Dim cCell As Range, Dcell As Range
For Each cCell In CredRNG
    If Not IsNumeric(cCell.Value) Then GoTo AlertTheUser
    Ctot = Ctot + cCell.Value
Next cCell
For Each Dcell In DebRNG
    If Not IsNumeric(Dcell.Value) Then GoTo AlertTheUser
    Dtot = Dtot + Dcell.Value
Next Dcell

If Ctot <> Dtot Then BalanceCredDeb = False
Exit Function
'Only numerics allowed
AlertTheUser:
MsgBox "Only Numbers can be entered as a debit or credit", vbCritical + vbOKOnly, "Illegal Character Detected"
End Function
