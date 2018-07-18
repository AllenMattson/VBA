VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm2"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   6630
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

Const GWL_STYLE = -16
Const WS_CAPTION = &HC00000

'UserForm position
Dim FormX As Double, FormY As Double


Private Sub RefreshIcon_Click()
    Call UpdateBox
End Sub

Private Sub UserForm_Initialize()
    Dim lngWindow As Long, lFrmHdl As Long
    lFrmHdl = FindWindowA(vbNullString, Me.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
End Sub

Private Sub TitleLabel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
'   Starting position when button is pressed
    If Button = 1 Then
        FormX = x
        FormY = Y
    End If
End Sub

Private Sub TitleLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
'   Move the form if the left button is pressed
    If Button = 1 Then
        Me.Left = Me.Left + (x - FormX)
        Me.Top = Me.Top + (Y - FormY)
    End If
End Sub

Private Sub UserForm_Activate()
    Call UpdateBox
End Sub

Private Sub CheckBoxRC_Click()
    Call UpdateBox
End Sub

Private Sub UpdateButton_Click()
    Call UpdateBox
End Sub

Private Sub CloseButton_Click()
    CheckingCells = False
    Set MyApp = Nothing
    Unload Me
End Sub

Sub UpdateBox()
    Dim DepCnt, DirDepCnt, PrecCnt, MaxField
    Dim Field(2, 1 To 12)
    Dim Labs(2, 1 To 12) As Control
    Dim i As Integer, j As Integer
    
    If TypeName(ActiveSheet) = "Worksheet" Then
        TitleLabel.Caption = "InfoBox for cell " & ActiveCell.Address(False, False) & " (" & ActiveCell.Address(True, True, xlR1C1) & ")"
    
'       Value of cell
        Field(2, 1) = ActiveCell.Value
        If Application.IsErr(ActiveCell) Then Field(2, 1) = ActiveCell.Text
        If Application.IsNA(ActiveCell) Then Field(2, 1) = ActiveCell.Text
    
'       Displayed in cell
        Field(2, 2) = ActiveCell.Text
        If Field(2, 2) = CStr(Field(2, 1)) Then Field(2, 2) = "(same)"
        If IsEmpty(ActiveCell) Then
            Field(2, 1) = "(empty)"
            Field(2, 2) = " "
        End If
    
'       Type of data
        Field(2, 3) = TypeName(ActiveCell.Value)
    
'       Number format
        Field(2, 4) = ActiveCell.NumberFormat
    
'       Formula
        If ActiveCell.HasFormula Then
            If CheckBoxRC Then
                Field(2, 5) = ActiveCell.FormulaR1C1
            Else
                Field(2, 5) = ActiveCell.Formula
            End If
        Else
            Field(2, 5) = "(none)"
        End If
    
'       Named cell
        Field(2, 6) = CellName(ActiveCell)
    
'       Protection status
        Field(2, 7) = "(not protected)"
        If ActiveCell.Locked Then Field(2, 7) = "Locked"
        If ActiveCell.FormulaHidden Then Field(2, 7) = "Hidden"
        If ActiveCell.Locked And ActiveCell.FormulaHidden Then Field(2, 7) = "Locked, Hidden"
        
'       Cell comment
        Field(2, 8) = CellComment(ActiveCell)
        
'       Dependent and precedent cells
        DepCnt = DependentCount(ActiveCell)
        If Not ActiveSheet.ProtectContents Then
            Select Case DepCnt
                Case 0
                    Field(2, 9) = "The cell is not used in any formulas."
                    Field(2, 10) = "The cell is not used in any formulas."
                Case Else
                    DirDepCnt = DirectDependentCount(ActiveCell)
                    Field(2, 9) = DepCnt
                    Field(2, 10) = DirDepCnt
            End Select

            If ActiveCell.HasFormula Then
                PrecCnt = PrecedentCount(ActiveCell)
                Select Case PrecCnt
                    Case 0
                        Field(2, 11) = "The cell does not use any other cells."
                        Field(2, 12) = "The cell does not use any other cells."
                    Case Else
                        Field(2, 11) = PrecedentCount(ActiveCell)
                        Field(2, 12) = DirectPrecedentCount(ActiveCell)
                End Select
            Else
                Field(2, 11) = "N/A"
                Field(2, 12) = "N/A"
            End If
        Else
            Field(2, 9) = "(unknown - protected sheet)"
            Field(2, 10) = "(unknown - protected sheet)"
            Field(2, 11) = "(unknown - protected sheet)"
            Field(2, 12) = "(unknown - protected sheet)"
        End If
    
    Else ' not a worksheet
        For i = 1 To 12
            Field(2, i) = "N/A"
        Next i
    End If

'   Create object variables
    Set Labs(1, 1) = Controls("Label1")
    Set Labs(1, 2) = Controls("Label3")
    Set Labs(1, 3) = Controls("Label5")
    Set Labs(1, 4) = Controls("Label7")
    Set Labs(1, 5) = Controls("Label9")
    Set Labs(1, 6) = Controls("Label11")
    Set Labs(1, 7) = Controls("Label13")
    Set Labs(1, 8) = Controls("Label15")
    Set Labs(1, 9) = Controls("Label17")
    Set Labs(1, 10) = Controls("Label19")
    Set Labs(1, 11) = Controls("Label21")
    Set Labs(1, 12) = Controls("Label23")
    
    Set Labs(2, 1) = Controls("Label2")
    Set Labs(2, 2) = Controls("Label4")
    Set Labs(2, 3) = Controls("Label6")
    Set Labs(2, 4) = Controls("Label8")
    Set Labs(2, 5) = Controls("Label10")
    Set Labs(2, 6) = Controls("Label12")
    Set Labs(2, 7) = Controls("Label14")
    Set Labs(2, 8) = Controls("Label16")
    Set Labs(2, 9) = Controls("Label18")
    Set Labs(2, 10) = Controls("Label20")
    Set Labs(2, 11) = Controls("Label22")
    Set Labs(2, 12) = Controls("Label24")

    
'   Labels for the fields
    Field(1, 1) = "Value:"
    Field(1, 2) = "Displayed As:"
    Field(1, 3) = "Cell Type:"
    Field(1, 4) = "Number Format:"
    Field(1, 5) = "Formula:"
    Field(1, 6) = "Name:"
    Field(1, 7) = "Protection:"
    Field(1, 8) = "Cell Comment:"
    Field(1, 9) = "Dependent Cells:"
    Field(1, 10) = "Dir Dependents:"
    Field(1, 11) = "Precedent Cells:"
    Field(1, 12) = "Dir Precedents:"

'   If sheet protected, use only the first 7
    If ActiveSheet.ProtectContents Then
        MaxField = 7
        Labs(1, 8).Visible = False
        Labs(1, 9).Visible = False
        Labs(1, 10).Visible = False
        Labs(1, 11).Visible = False
        Labs(2, 8).Visible = False
        Labs(2, 9).Visible = False
        Labs(2, 10).Visible = False
        Labs(2, 11).Visible = False
        Labs(2, 12).Visible = False
    Else
        MaxField = 12
        Labs(1, 8).Visible = True
        Labs(1, 9).Visible = True
        Labs(1, 10).Visible = True
        Labs(1, 11).Visible = True
        Labs(2, 8).Visible = True
        Labs(2, 8).Visible = True
        Labs(2, 10).Visible = True
        Labs(2, 11).Visible = True
        Labs(2, 12).Visible = True
    End If
    
'   Transfer the data to labels
    For i = 1 To 2
        For j = 1 To MaxField
            Labs(i, j) = Field(i, j)
        Next j
    Next i

'   Adjust the width and height
    For j = 1 To MaxField
        Labs(1, j).Left = 10
        Labs(1, j).Width = 68
        
        Labs(2, j).Left = 84
        Labs(2, j).Width = 220
        Labs(2, j).AutoSize = False
        Labs(2, j).Width = 205
        Labs(2, j).AutoSize = True
        Labs(2, j).Width = 220
    Next j

'   Adjust the spacing
    For j = 2 To 12
        Labs(2, j).Top = Labs(2, j - 1).Top + Labs(2, j - 1).Height + 5
        Labs(1, j).Top = Labs(2, j - 1).Top + Labs(2, j - 1).Height + 5
    Next j
    
'   Adjust the Userform height
    Frame2.ScrollHeight = Labs(2, MaxField).Top + Labs(2, MaxField).Height
    
    CheckingCells = False
End Sub

Private Function CellName(c)
'   Returns True if the cell has a name
    CellName = "(none)"
    On Error Resume Next
    CellName = c.Name.Name
End Function

Private Function CellComment(c)
'   Returns the cell comment, or "(none)"
    CellComment = "(none)"
    On Error Resume Next
    CellComment = c.Comment.Text
End Function

Function DependentCount(cell) As Variant
'   Returns the number of dependent cells
'   Same sheet only!
    On Error Resume Next
    CheckingCells = True
    DependentCount = cell.Dependents.Count
    If Err <> 0 Then DependentCount = 0
    On Error GoTo 0
End Function

Function DirectDependentCount(cell) As Variant
'   Returns the number of direct dependents
'   Same sheet only!
    On Error Resume Next
    CheckingCells = True
    DirectDependentCount = cell.DirectDependents.Count
    If Err <> 0 Then DirectDependentCount = 0
    On Error GoTo 0
End Function

Function PrecedentCount(cell) As Variant
'   Returns the number of cell precenents
'   Same sheet only!
    Dim x
    On Error Resume Next
    CheckingCells = True
    PrecedentCount = cell.Precedents.Count
    If Err <> 0 Then
        PrecedentCount = 0
    End If
    On Error GoTo 0
    x = 1
End Function

Function DirectPrecedentCount(cell) As Variant
'   Returns the number of direct precedents
'   Same sheet only
    Dim x
    On Error Resume Next
    CheckingCells = True
    DirectPrecedentCount = cell.DirectPrecedents.Count
    If Err <> 0 Then
        DirectPrecedentCount = 0
    End If
    On Error GoTo 0
    x = 1
End Function


