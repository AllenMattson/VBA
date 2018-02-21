VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Array Sorting Demo"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub OKButton_Click()
    Dim Array1() As Long, Array2() As Long, Array3() As Long, Array4() As Long
    Dim i As Long
    Dim Elements As Long, RElement As Long, Temp As Long
    Dim Time1 As Date, Time2 As Date
    Dim Time3 As Date, Time4 As Date
    Dim Time5 As Date, Time6 As Date
    Dim Time7 As Date, Time8 As Date
    Dim Msg As String
    Dim r As Long
        
    lblTime1.Caption = ""
    lblTime2.Caption = ""
    lblTime3.Caption = ""
    LblTime4.Caption = ""
    
'   Valid userform entries
    If Not IsNumeric(tbElements.Value) Then
        MsgBox "Invalid entry (No. elements)", vbInformation
        tbElements.SetFocus
        Exit Sub
    End If
    Elements = Val(tbElements.Value)
    
    If Elements < 1 Then
        MsgBox "Invalid entry (No. elements)", vbInformation
        tbElements.SetFocus
        Exit Sub
    End If
    
'   Build four identical arrays
    lblCurrentSort = "Building the array..."
    Me.Repaint
    Randomize
    ReDim Array1(1 To Elements, 0)
    ReDim Array2(1 To Elements)
    ReDim Array3(1 To Elements)
    ReDim Array4(1 To Elements)
    For i = 1 To Elements
        Array1(i, 0) = CLng(Rnd * 100000)
        Array2(i) = Array1(i, 0)
        Array3(i) = Array1(i, 0)
        Array4(i) = Array1(i, 0)
    Next i

'   Worksheet sort
    If CheckBox1 Then
        lblCurrentSort = "Performing Worksheet Sort..."
        Me.Repaint
        Time1 = Timer
        Call WorksheetSort(Array1)
        Time2 = Timer
        lblTime1.Caption = Format(Time2 - Time1, "00.00") & " sec."
        Me.Repaint
    End If

'   VBA Bubble Sort
    If CheckBox2 Then
        lblCurrentSort = "Performing Bubble Sort..."
        Me.Repaint
        Time1 = Timer
        Call BubbleSort(Array2)
        Time2 = Timer
        lblTime2.Caption = Format(Time2 - Time1, "00.00") & " sec."
        Me.Repaint
    End If

'   VBA Quick Sort
    If CheckBox3 Then
        lblCurrentSort = "Performing Quick Sort..."
        Me.Repaint
        Time1 = Timer
        Call Quicksort(Array3, LBound(Array2), UBound(Array2))
        Time2 = Timer
        lblTime3.Caption = Format(Time2 - Time1, "00.00") & " sec."
        Me.Repaint
    End If

'   VBA Counting Sort
    If CheckBox4 Then
        lblCurrentSort = "Performing Counting Sort..."
        Me.Repaint
        Time1 = Timer
        Call CountingSort(Array4)
        Time2 = Timer
        LblTime4.Caption = Format(Time2 - Time1, "00.00") & " sec."
        Me.Repaint
    End If

'   write the sorted data
'   Arrays that exceed 65,536 elements won't get written because
'   of a limitation by Transpose
    Worksheets("data").Activate
    Cells.Clear
    On Error Resume Next
    Range(Cells(1, 1), Cells(1, 4)) = Array("Worksheet", "Bubble", "Quck", "Counting")
    Range(Cells(2, 1), Cells(UBound(Array1) + 1, 1)) = Array1
    Range(Cells(2, 2), Cells(UBound(Array2) + 1, 2)) = Application.Transpose(Array2)
    Range(Cells(2, 3), Cells(UBound(Array3) + 1, 3)) = Application.Transpose(Array3)
    Range(Cells(2, 4), Cells(UBound(Array4) + 1, 4)) = Application.Transpose(Array4)
    'If Err.Value <> 0 Then Cells(2, 1) = "Too much data"
    lblCurrentSort = "Finished."
End Sub

Private Sub CancelButton_Click()
    Unload Me
End Sub


