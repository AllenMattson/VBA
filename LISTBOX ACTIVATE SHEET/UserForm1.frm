VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "GoTo Sheet"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OriginalSheet As Object

Private Sub UserForm_Initialize()
    Dim SheetData() As String
    Dim ShtCnt As Integer
    Dim ShtNum As Integer
    Dim Sht As Object
    Dim ListPos As Integer
    
    Set OriginalSheet = ActiveSheet
    ShtCnt = ActiveWorkbook.Sheets.Count
    ReDim SheetData(1 To ShtCnt, 1 To 4)
    ShtNum = 1
    For Each Sht In ActiveWorkbook.Sheets
        If Sht.Name = ActiveSheet.Name Then _
          ListPos = ShtNum - 1
        SheetData(ShtNum, 1) = Sht.Name
        Select Case TypeName(Sht)
            Case "Worksheet"
                SheetData(ShtNum, 2) = "Sheet"
                SheetData(ShtNum, 3) = _
                  Application.CountA(Sht.Cells)
            Case "Chart"
                SheetData(ShtNum, 2) = "Chart"
                SheetData(ShtNum, 3) = "N/A"
            Case "DialogSheet"
                SheetData(ShtNum, 2) = "Dialog"
                SheetData(ShtNum, 3) = "N/A"
        End Select
        If Sht.Visible Then
            SheetData(ShtNum, 4) = "True"
        Else
            SheetData(ShtNum, 4) = "False"
        End If
        ShtNum = ShtNum + 1
    Next Sht
    With ListBox1
        .ColumnWidths = "100 pt;30 pt;40 pt;50 pt"
        .List = SheetData
        .ListIndex = ListPos
    End With
End Sub

Private Sub CancelButton_Click()
    OriginalSheet.Activate
    Unload Me
End Sub

Private Sub cbPreview_Click()
    If cbPreview Then Sheets(ListBox1.Value).Activate
End Sub

Private Sub ListBox1_Click()
    If cbPreview Then _
        Sheets(ListBox1.Value).Activate
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call OKButton_Click
End Sub

Private Sub OKButton_Click()
    Dim UserSheet As Object
    Set UserSheet = Sheets(ListBox1.Value)
    If UserSheet.Visible Then
        UserSheet.Activate
    Else
        If MsgBox("Unhide sheet?", _
          vbQuestion + vbYesNoCancel) = vbYes Then
            UserSheet.Visible = True
            UserSheet.Activate
        Else
            OriginalSheet.Activate
        End If
    End If
    Unload Me
End Sub

