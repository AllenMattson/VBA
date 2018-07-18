VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Salaries 
   Caption         =   "Employees and Salaries"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   OleObjectBlob   =   "Salaries.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Salaries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
' Global declarations
Dim emp As New CEmployee
Dim CEmployees As New Collection
Dim index As Integer
Dim ws As Worksheet
Dim extract As String
Dim cell As Range
Dim lastRow As Integer
Dim empLoc As Integer
Dim startRow As Integer
Dim endRow As Integer
Dim choice As Integer
Dim amount As Long

Private Sub cmdEmployeeList_Click()
    lboxPeople.Clear
    For Each emp In CEmployees
    lboxPeople.AddItem emp.ID & ", " & _
        emp.LastName & ", " & emp.FirstName & ", $" & _
        Format(emp.Salary, "0.00")
    Next emp
End Sub

Private Sub cmdUpdate_Click()
    If optHighlighted = False And optAll = False Then
        MsgBox "Click the 'Highlighted Employee' or " _
            & " 'All Employees' option button."
        Exit Sub
    End If
    If Not IsNumeric(txtRaise) Then
        MsgBox "This field requires a number."
        txtRaise.SetFocus
        Exit Sub
    End If
    If optHighlighted = True And _
        lboxPeople.ListIndex = -1 Then
            MsgBox "Click the name of the employee."
            Exit Sub
    End If
    If lboxPeople.ListIndex <> -1 And _
        optHighlighted = True And _
        optAmount.Value = True And _
        txtRaise.Value <> "" Then
            extract = CEmployees.Item(lboxPeople.ListIndex + 1).ID
            MsgBox extract
            Call FindId
            MsgBox empLoc
            choice = 2
            amount = txtRaise
            CEmployees.Item(lboxPeople.ListIndex + 1).Salary = _
                emp.CalcNewSalary(choice, _
            CEmployees.Item(lboxPeople.ListIndex + 1).Salary, amount)
            Range("D" & empLoc).Formula = CEmployees. _
                Item(lboxPeople.ListIndex + 1).Salary
            cmdEmployeeList.Value = True
    ElseIf lboxPeople.ListIndex <> -1 And _
            optHighlighted = True And _
            optPercent.Value = True And _
            txtRaise.Value <> "" Then
                extract = CEmployees.Item(lboxPeople.ListIndex + 1).ID
                MsgBox extract
                Call FindId
                MsgBox empLoc
                CEmployees.Item(lboxPeople.ListIndex + 1).Salary = _
                    CEmployees.Item(lboxPeople.ListIndex + 1).Salary + _
                    (CEmployees.Item(lboxPeople.ListIndex + 1).Salary * _
                    txtRaise / 100)
                Range("D" & empLoc).Formula = CEmployees. _
                    Item(lboxPeople.ListIndex + 1).Salary
                cmdEmployeeList.Value = True
    ElseIf optAll = True And _
            optPercent.Value = True And _
            txtRaise.Value <> "" Then
                For Each emp In CEmployees
                    emp.Salary = emp.Salary + ((emp.Salary * txtRaise) / 100)
                    extract = emp.ID
                    MsgBox extract
                    Call FindId
                    MsgBox empLoc
                    Range("D" & empLoc).Formula = emp.Salary
                Next emp
                cmdEmployeeList.Value = True
    ElseIf optAll = True And _
            optAmount.Value = True And _
            txtRaise.Value <> "" Then
                For Each emp In CEmployees
                    emp.Salary = emp.Salary + txtRaise
                    extract = emp.ID
                    MsgBox extract
                    Call FindId
                    MsgBox empLoc
                    Range("D" & empLoc).Formula = emp.Salary
                Next emp
                cmdEmployeeList.Value = True
    Else
        MsgBox "Enter data or select an option."
    End If
End Sub

Private Sub cmdSave_Click()
    If txtLastName.Value = "" Or txtFirstName.Value = "" Or _
      txtSalary.Value = "" Then
        MsgBox "Enter Last Name, First Name and Salary."
        txtLastName.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtSalary) Then
        MsgBox "You must enter a value for the Salary."
    txtSalary.SetFocus
        Exit Sub
    End If
    If txtSalary < 0 Then
        MsgBox "Salary cannot be a negative number."
        Exit Sub
    End If
    Worksheets("Salaries").Select
    index = ActiveSheet.UsedRange.Rows.Count + 1
    lboxPeople.Enabled = True
    'set and enter data into the CEmployees collection
    With emp
        Cells(index, 1).Formula = emp.ID
        .LastName = txtLastName
        Cells(index, 2).Formula = emp.LastName
        .FirstName = txtFirstName
        Cells(index, 3).Formula = emp.FirstName
        .Salary = CCur(txtSalary)
        If .Salary = 0 Then Exit Sub
        Cells(index, 4).Formula = emp.Salary
        CEmployees.Add emp
    End With
    'delete data from text boxes
    txtLastName = ""
    txtFirstName = ""
    txtSalary = ""
    'enable hidden controls
    cmdEmployeeList.Value = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    Frame1.Enabled = True
    txtRaise.Enabled = True
    optPercent.Enabled = True
    optAmount.Enabled = True
    Frame2.Enabled = True
    optHighlighted.Enabled = True
    optAll.Enabled = True
    txtLastName.SetFocus
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
' make sure an employee is highlighted in the
' list box control
 
    If lboxPeople.ListIndex > -1 Then
        MsgBox "Selected item number: " & lboxPeople.ListIndex
        extract = CEmployees.Item(lboxPeople.ListIndex + 1).ID
        MsgBox extract
        
        Call FindId
        MsgBox empLoc
        Range("A" & empLoc).Delete (3)
        
        MsgBox "There are " & CEmployees.Count & _
            " items in the CEmployees collection. "
        CEmployees.Remove lboxPeople.ListIndex + 1
        MsgBox "The CEmployees collection has now " & CEmployees.Count & " items."
        cmdEmployeeList.Value = True
            If CEmployees.Count = 0 Then
                Call UserForm_Initialize
            End If
    Else
         MsgBox "Click the item you want to remove."
    End If
End Sub

Private Sub UserForm_Initialize()
    txtLastName.SetFocus
    cmdEmployeeList.Visible = False
    lboxPeople.Enabled = False
    Frame1.Enabled = False
    
    txtRaise.Value = ""
    optPercent.Value = False
    optAmount.Value = False
    
    txtRaise.Enabled = False
    optPercent.Enabled = False
    optAmount.Enabled = False
    
    Frame2.Enabled = False
    optHighlighted.Enabled = False
    optAll.Enabled = False
    
    cmdUpdate.Enabled = False
    cmdDelete.Enabled = False
End Sub

Private Function FindId()
    Set ws = ActiveWorkbook.Sheets("Salaries")
    startRow = ActiveSheet.UsedRange.Rows.Count + 1 - CEmployees.Count
    endRow = ActiveSheet.UsedRange.Rows.Count
    For Each cell In ws.Range(Cells(startRow, 1), Cells(endRow, 1))
        If cell.Value = extract Then
          empLoc = cell.Row
          FindId = empLoc
            Exit Function
        End If
    Next
End Function



