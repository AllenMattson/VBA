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

Private OriginalSheet As Object

Private Sub UserForm_Initialize()
    Dim SheetData() As String
    Dim ShtCnt As Long
    Dim ShtNum As Long
    Dim Sht As Object
    Dim ListPos As Long
    
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
    With Me.lbxSheets
        .ColumnWidths = "100 pt;30 pt;40 pt;50 pt"
        .List = SheetData
        .ListIndex = ListPos
    End With
End Sub

Private Sub cmdCancel_Click()
    OriginalSheet.Activate
    Unload Me
End Sub

Private Sub chkPreview_Click()
    If chkPreview Then Sheets(Me.lbxSheets.Value).Activate
End Sub

Private Sub lbxSheets_Click()
    If chkPreview Then _
        Sheets(Me.lbxSheets.Value).Activate
End Sub

Private Sub lbxSheets_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOK_Click
End Sub

Private Sub cmdOK_Click()
    Dim UserSheet As Object
    Set UserSheet = Sheets(Me.lbxSheets.Value)
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

