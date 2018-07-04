Attribute VB_Name = "JemBalanceButton"
Option Explicit
Private Ctot As Double 'credit total
Private Dtot As Double 'debit total
Private Btot As Currency
Public Sub JemBalanceButton()
Ctot = 0
Dtot = 0
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''This macro changes colors depending'''''''''''''''
'''''''''''''if the credits and debits balance''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'On Error GoTo NoValueFound
Dim ObjSh As Shape 'initialize a shape variable for text boxes
ActiveCell.Offset(4, 0).Range("A1").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlDown)).Select
   
Dim Credits As Range, Debits As Range, TotalRange As Range
Dim Clr As Long
Dim Dlr As Long
With ActiveSheet
    Set Credits = Range("I6:I999").SpecialCells(xlCellTypeConstants, 23)
    Set Debits = .Range("J6:J999").SpecialCells(xlCellTypeConstants, 23)
End With
Dim TotalsBoxes As Range
Set TotalsBoxes = Range("C1, E1, H1")
TotalsBoxes.ClearContents
Set TotalRange = Range(Credits, Debits)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Pass all values from credit range and debit range
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If BalanceCredDeb(Credits, Debits) = False Then
    With TotalRange.Font
        .Color = vbBlue
    End With
    With TotalsBoxes
        .Interior.Color = vbBlue
        .Font.Color = vbWhite
        .Font.Bold = False
        .Font.Size = 16
    End With
Else
    With TotalRange.Font
        .Color = vbBlack
        .Bold = True
    End With
    With TotalsBoxes
        .Interior.Color = vbWhite
        .Font.Color = vbBlack
        .Font.Bold = True
        .Font.Size = 16
    End With
End If

'Put entry totals in their corresponding cells
'format the ranges

Btot = Ctot - Dtot
With ActiveSheet
    .Range("C1").Formula = "=Round(" & Ctot & ", 2)"
    .Range("E1").Formula = "=Round(" & Dtot & ", 2)"
    .Range("H1").value = Ctot - Dtot
End With
Exit Sub

'Set everything to normal, alert user if no entry values found in debits or credits
NoValueFound:
If Err.Number = 1004 Then MsgBox "No Credit or Debit Found" & vbNewLine & vbNewLine & "Error: " & Err.Number & vbNewLine & Err.Description, vbInformation + vbOKOnly, "Credit or Debit Needed"
With ActiveSheet.Range("J9:J10000").Font
    .Color = vbBlack
    .Bold = False
End With
With Range("I9:I10000").Font
    .Color = vbBlack
    .Bold = False
End With
With TotalsBoxes
    .Interior.Color = vbWhite
    .Font.Color = vbBlack
    .Font.Bold = False
    .Font.Size = 12
    .value = ""
End With
End Sub
Private Function BalanceCredDeb(CredRNG As Range, DebRNG As Range) As Boolean
'Add numeric values in range and make sure they balance
BalanceCredDeb = True
Dim cCell As Range, Dcell As Range
For Each cCell In CredRNG
    If Not IsNumeric(cCell.value) Then GoTo AlertTheUser
    Ctot = Ctot + cCell.value
Next cCell
For Each Dcell In DebRNG
    If Not IsNumeric(Dcell.value) Then GoTo AlertTheUser
    Dtot = Dtot + Dcell.value
Next Dcell

If Ctot <> Dtot Then BalanceCredDeb = False
Exit Function
'Only numerics allowed
AlertTheUser:
MsgBox "Only Numbers can be entered as a debit or credit", vbCritical + vbOKOnly, "Illegal Character Detected"
End Function
