VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Students 
   Caption         =   "Students and Exams"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8460
   OleObjectBlob   =   "Students.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Students"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Global Declarations
Dim r As Integer
Dim indexPlus As Integer
Dim YesNo As Integer
Dim startRow As Integer
Dim endRow As Integer


Private Sub cboxGrade_Click()
    YesNo = MsgBox("Enter the grade (" & cboxGrade.Value & _
        ") in the worksheet?", vbYesNo, "Enter/Modify Grade")
    If YesNo = 6 Then
       Me.lblGrade.Caption = cboxGrade.Value
       Select Case TabStrip1.Value
        Case 0
          Range("F" & indexPlus).Value = Me.lblGrade.Caption
        Case 1
            Range("H" & indexPlus).Value = Me.lblGrade.Caption
        Case 2
            Range("J" & indexPlus).Value = Me.lblGrade.Caption
        Case 3
            Range("L" & indexPlus).Value = Me.lblGrade.Caption
        End Select
        cboxGrade.Value = ""
    End If
End Sub



Private Sub cmdClose_Click()
    Unload Me
    Set Students = Nothing
End Sub

Private Sub cmdAddStudent_Click()
    If IsFilled = False Then
        MsgBox "Please enter all the data " & _
            "or click Close to exit."
        Me.txtSSN.SetFocus
        Exit Sub
    End If
    
    With Me.txtSSN
        If Not IsNumeric(.Text) Or Len(.Text) <> 9 Then
            MsgBox "Please enter a 9-digit number with no dashes.", _
                vbExclamation + vbOKOnly, "Invalid SSN"
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
            Exit Sub
        End If
    End With
    
    If Me.optDataEntry.Value = True Then
        Me.Hide
        ActiveWorkbook.Sheets("Students and Exams").Select
        r = ActiveSheet.UsedRange.Rows.Count
        r = r + 1
        Range("A" & r).Value = Me.txtSSN.Text
        Range("B" & r).Value = Me.txtLast.Text
        Range("C" & r).Value = Me.txtFirst.Text
        Range("D" & r).Value = Me.cboxYear.Text
        Range("F" & r).Value = Me.cboxMajor.Text
        
        With Me
            .txtSSN.Text = ""
            .txtLast.Text = ""
            .txtFirst.Text = ""
            .cboxYear.Text = ""
            .cboxMajor.Text = ""
            .txtSSN.SetFocus
            .optViewData.Enabled = True
            ' redisplay the form
            .Show
        End With
    End If
End Sub


Private Sub DTPicker1_Change()
    Me.lblDate.Caption = Me.DTPicker1.Value
End Sub

Private Sub DTPicker1_CloseUp()
    YesNo = MsgBox("Enter the selected date " & DTPicker1.Value & _
        " in the worksheet?", vbYesNo, "Modify Exam Date")
    If YesNo = 6 Then
        Me.lblDate.Caption = Format(DTPicker1.Value, "mm/dd/yyyy")
        Select Case TabStrip1.Value
            Case 0
                Range("G" & indexPlus).Value = Me.lblDate.Caption
            Case 1
                Range("I" & indexPlus).Value = Me.lblDate.Caption
            Case 2
                Range("K" & indexPlus).Value = Me.lblDate.Caption
            Case 3
                Range("M" & indexPlus).Value = Me.lblDate.Caption
        End Select
    End If
End Sub

Private Sub DTPicker1_Enter()
    DTPicker1.maxDate = Date
    SendKeys "{F4}"
End Sub

Private Sub DTPicker1_MouseUp(ByVal Button As Integer, _
    ByVal Shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, _
    ByVal y As stdole.OLE_YPOS_PIXELS)
    
    SendKeys "{F4}"
End Sub

Private Sub lboxStudents_Change()
    indexPlus = lboxStudents.ListIndex + 3
 
    With ActiveWorkbook.Worksheets("Students and Exams")
        Me.txtSSN.Text = Range("A" & indexPlus).Value
        Me.txtLast.Text = Range("B" & indexPlus).Value
        Me.txtFirst.Text = Range("C" & indexPlus).Value
        Me.cboxYear.Text = Range("D" & indexPlus).Value
        Me.cboxMajor.Text = Range("F" & indexPlus).Value
    
     Call TabStrip1_Change
     Me.MultiPage1(1).Enabled = True
    End With
End Sub


Private Sub MultiPage1_Change()
    If MultiPage1.Value = 1 Then
        Me.lblWho.Caption = Me.txtLast.Value & ", " _
            & Me.txtFirst.Value
        Me.TabStrip1.Value = 0
    End If
End Sub

Private Sub optDataEntry_Click()
    cmdAddStudent.Visible = True
    lblNames.Visible = False
    lboxStudents.Visible = False
        
    Me.MultiPage1(1).Enabled = False
        If lboxStudents.RowSource <> "" Then
            Me.txtSSN.Text = ""
            Me.txtLast.Text = ""
            Me.txtFirst.Text = ""
            Me.cboxYear.Text = ""
            Me.cboxMajor.Text = ""
            Me.txtSSN.SetFocus
        End If
    Me.txtSSN.SetFocus
End Sub

Private Sub optViewData_Click()
    startRow = 3
    endRow = ActiveSheet.UsedRange.Rows.Count
    If endRow > 2 Then
        lblNames.Visible = True
        lboxStudents.Visible = True
        lboxStudents.RowSource = _
          Range(Cells(startRow, 2), Cells(endRow, 3)).Address
        lboxStudents.ListIndex = 0
        cmdAddStudent.Visible = False
   End If
End Sub


Private Sub TabStrip1_Change()
    indexPlus = lboxStudents.ListIndex + 3

    With ActiveWorkbook.Worksheets("Students and Exams")
        Select Case TabStrip1.Value
            Case 0  ' English
                Me.lblGrade.Caption = Range("F" & indexPlus).Value
                Me.lblDate.Caption = Range("G" & indexPlus).Value
            Case 1 'French
                Me.lblGrade.Caption = Range("H" & indexPlus).Value
                Me.lblDate.Caption = Range("I" & indexPlus).Value
            Case 2  'Math
                Me.lblGrade.Caption = Range("J" & indexPlus).Value
                Me.lblDate.Caption = Range("K" & indexPlus).Value
            Case 3 'Physics
                Me.lblGrade.Caption = Range("L" & indexPlus).Value
                Me.lblDate.Caption = Range("M" & indexPlus).Value
        End Select
    End With
End Sub







Private Sub UserForm_Initialize()
    With Me.MultiPage1(0)
       optDataEntry.Value = True
        optViewData.Value = False
    End With
        
    If ActiveSheet.UsedRange.Rows.Count > 2 Then
        optViewData.Enabled = True
    Else
        optViewData.Enabled = False
        lboxStudents.Visible = False
    End If
    lblNames.Visible = False
         
    With Me.cboxYear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
    End With
    
    With Me.cboxMajor
        .AddItem "English"
        .AddItem "Chemistry"
        .AddItem "Mathematics"
        .AddItem "Linguistics"
        .AddItem "Computer Science"
    End With
    
    With Me.cboxGrade
        .AddItem "A"
        .AddItem "B"
        .AddItem "C"
        .AddItem "D"
        .AddItem "F"
    End With
    
    Me.txtSSN.SetFocus
End Sub



