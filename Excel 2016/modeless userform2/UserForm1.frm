VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Cell Info Box"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const msNONE As String = "(none)"
Private Const msNODEPS As String = "The cell is not used in any formulas."
Private Const msNOPRECS As String = "The cell does not use any other cells."
Private Const msPROTECTED As String = "(unknown - protected sheet)"
Private Const msNA As String = "N/A"

Private Sub UserForm_Activate()
    UpdateBox
End Sub

Private Sub chkR1C1_Click()
    UpdateBox
End Sub

Private Sub cmdUpdate_Click()
    UpdateBox
End Sub

Private Sub cmdClose_Click()
    CheckingCells = False
    Set gclsEvent = Nothing
    Unload Me
End Sub

Sub UpdateBox()
    Dim DepCnt As Long, DirDepCnt As Long, PrecCnt As Long, MaxField As Long
    Dim i As Long
    Dim vaLabels As Variant
    Dim ctl As Control
    Dim ctlLabel As MSForms.Label
    Dim ctlData As MSForms.Label
    Dim ctlDataPrev As MSForms.Label
    
    Const sSAME As String = "(same)"
    Const sEMPTY As String = "(empty)"

    If TypeName(ActiveSheet) = "Worksheet" Then
        Me.Caption = "InfoBox for cell: " & ActiveCell.Address(False, False) & " (" & ActiveCell.Address(True, True, xlR1C1) & ")"
        
        'Fill in the labels
        vaLabels = Split("Value:|Displayed As:|Cell Type:|Number Format:|Formula:|Name:|Protection:|Cell Comment:|Dependent Cells:|Dir Dependents:|Precedent Cells:|Dir Precedents:", "|")
        For i = 1 To 12
            Me.Controls("Label1" & Format(i, "00")).Caption = vaLabels(i - 1)
        Next i
        
'       Value and Text of cell
        If Application.IsErr(ActiveCell) Or Application.IsNA(ActiveCell) Then
            Me.lblValue.Caption = ActiveCell.Text
            Me.lblAsDisplayed.Caption = ActiveCell.Text
        ElseIf IsEmpty(ActiveCell.Value) Then
            Me.lblValue.Caption = sEMPTY
            Me.lblAsDisplayed.Caption = sEMPTY
        Else
            Me.lblValue.Caption = ActiveCell.Value
            If ActiveCell.Value = ActiveCell.Text Then
                Me.lblAsDisplayed.Caption = sSAME
            Else
                Me.lblAsDisplayed.Caption = ActiveCell.Text
            End If
        End If
    
'       Type of data
        Me.lblDataType.Caption = TypeName(ActiveCell.Value)
    
'       Number format
        Me.lblFormat.Caption = ActiveCell.NumberFormat
    
'       Formula
        If ActiveCell.HasFormula Then
            If Me.chkR1C1.Value Then
                Me.lblFormula.Caption = ActiveCell.FormulaR1C1
            Else
                Me.lblFormula.Caption = ActiveCell.Formula
            End If
        Else
            Me.lblFormula.Caption = msNONE
        End If
    
'       Named cell
        Me.lblName.Caption = CellName(ActiveCell)
    
'       Protection status
        Me.lblProtected.Caption = Array("(unprotected)", "Locked", "Hidden", "Locked, Hidden")(Abs(CLng(ActiveCell.Locked) + 2 * CLng(ActiveCell.FormulaHidden)))
        
'       Cell comment
        Me.lblComment.Caption = CellComment(ActiveCell)
        
'       Dependent and precedent cells
        Me.lblDepCnt.Caption = DependentCount(ActiveCell, False)
        Me.lblDirDepCnt.Caption = DependentCount(ActiveCell, True)
        Me.lblPrecCnt.Caption = PrecedentCount(ActiveCell, False)
        Me.lblDirPrecCnt.Caption = PrecedentCount(ActiveCell, True)
    
    Else ' not a worksheet
        For Each ctl In Me.Controls
            If TypeName(ctl) = "Label" Then
                If Not ctl.Name Like "Label###" Then
                    ctl.Caption = msNA
                End If
            End If
        Next ctl
    End If
    
'   Hide labels below Protected for protected worksheets
    For Each ctl In Me.Controls
        If TypeName(ctl) = "Label" Then
            If ctl.Top > Me.lblProtected.Top Then
                ctl.Visible = Not ActiveSheet.ProtectContents
            End If
        End If
    Next ctl
    
    'Set the heigths and widths
    For i = 1 To 12
        Set ctlLabel = ControlByTag("1" & Format(i, "00"))
        Set ctlData = ControlByTag("2" & Format(i, "00"))
        
        ctlLabel.Left = 10
        ctlLabel.Width = 68
        ctlData.Left = 84
        ctlData.Width = 220
        ctlData.AutoSize = False
        ctlData.Width = 205
        ctlData.AutoSize = True
        ctlData.Width = 220
        
        If i > 1 Then
            Set ctlDataPrev = ControlByTag("2" & Format(i - 1, "00"))
            ctlData.Top = ctlDataPrev.Top + ctlDataPrev.Height + 5
        End If
        
        ctlLabel.Top = ctlData.Top
    Next i

'   Adjust the Userform height
    If Me.lblDirPrecCnt.Visible Then
        Me.Height = Me.lblDirPrecCnt.Top + Me.lblDirPrecCnt.Height + 35
    Else
        Me.Height = Me.lblProtected.Top + Me.lblProtected.Height + 35
    End If
    
    CheckingCells = False
End Sub

Private Function CellName(c As Range) As String
'   Returns True if the cell has a name
    CellName = msNONE
    On Error Resume Next
    CellName = c.Name.Name
End Function

Private Function CellComment(c As Range) As String
'   Returns the cell comment, or "(none)"
    CellComment = msNONE
    On Error Resume Next
    CellComment = c.Comment.Text
End Function

Function DependentCount(cell As Range, bDirect As Boolean) As String
'   Returns the number of dependent cells
'   Same sheet only!
    Dim lCnt As Long
    
    On Error Resume Next
        CheckingCells = True
        If bDirect Then
            lCnt = cell.DirectDependents.Count
        Else
            lCnt = cell.Dependents.Count
        End If
        If Err <> 0 Then lCnt = 0
    On Error GoTo 0
    
    If Not cell.Parent.ProtectContents Then
        If lCnt = 0 Then
            DependentCount = msNODEPS
        Else
            DependentCount = lCnt
        End If
    Else
        DependentCount = msPROTECTED
    End If

End Function

Function PrecedentCount(cell As Range, bDirect As Boolean) As String
'   Returns the number of cell precenents
'   Same sheet only!
    
    Dim lCnt As Long
    
    On Error Resume Next
        CheckingCells = True
        If bDirect Then
            lCnt = cell.DirectPrecedents.Count
        Else
            lCnt = cell.Precedents.Count
        End If
        If Err <> 0 Then PrecedentCount = 0
    On Error GoTo 0
    
    If Not cell.Parent.ProtectContents Then
        If cell.HasFormula Then
            If lCnt = 0 Then
                PrecedentCount = msNOPRECS
            Else
                PrecedentCount = lCnt
            End If
        Else
            PrecedentCount = msNA
        End If
    Else
        PrecedentCount = msPROTECTED
    End If
    
End Function

Private Function ControlByTag(ByVal sTag As String) As MSForms.Label
    
    Dim ctl As Control
    
    For Each ctl In Me.Controls
        If ctl.Tag = sTag Then
            Set ControlByTag = ctl
            Exit For
        End If
    Next ctl
    
End Function
