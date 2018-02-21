Attribute VB_Name = "Module1"
Option Explicit

Sub AddSubmenu()
'   Adds a submenu to the six shortcut menus
    Dim Bar As CommandBar
    Dim NewMenu As CommandBarControl
    Dim NewSubmenu As CommandBarButton
    
    DeleteSubmenu
    Set Bar = CommandBars("Cell")
'   Add submenu
    Set NewMenu = Bar.Controls.Add _
        (Type:=msoControlPopup, _
         temporary:=True)
    NewMenu.Caption = "Ch&ange Case"
    NewMenu.BeginGroup = True
'   Add first submenu item
    Set NewSubmenu = NewMenu.Controls.Add _
      (Type:=msoControlButton)
    With NewSubmenu
        .FaceId = 38
        .Caption = "&Upper Case"
        .OnAction = "MakeUpperCase"
    End With
'   Add second submenu item
    Set NewSubmenu = NewMenu.Controls.Add _
      (Type:=msoControlButton)
    With NewSubmenu
        .FaceId = 40
        .Caption = "&Lower Case"
        .OnAction = "MakeLowerCase"
    End With
'   Add third submenu item
    Set NewSubmenu = NewMenu.Controls.Add _
      (Type:=msoControlButton)
    With NewSubmenu
        .FaceId = 476
        .Caption = "&Proper Case"
        .OnAction = "MakeProperCase"
    End With
End Sub

Sub DeleteSubmenu()
    On Error Resume Next
    CommandBars("Cell").Controls("Cha&nge Case").Delete
End Sub


Sub MakeUpperCase()
    Dim WorkRange As Range
    Dim cell As Range
'   Process only text cells, no formulas
    On Error Resume Next
    If Selection.Count = 1 Then
        Set WorkRange = ActiveCell
    Else
        Set WorkRange = Selection.SpecialCells _
       (xlCellTypeConstants, xlCellTypeConstants)
    End If
'   Upper case
    Application.ScreenUpdating = False
    For Each cell In WorkRange
        If Not cell.HasFormula Then cell.Value = UCase(cell.Value)
    Next cell
End Sub

Sub MakeLowerCase()
    Dim WorkRange As Range
    Dim cell As Range
'   Process only text cells, no formulas
    On Error Resume Next
    If Selection.Count = 1 Then
        Set WorkRange = ActiveCell
    Else
        Set WorkRange = Selection.SpecialCells _
       (xlCellTypeConstants, xlCellTypeConstants)
    End If
'   Lower case
    Application.ScreenUpdating = False
    For Each cell In WorkRange
        If Not cell.HasFormula Then cell.Value = LCase(cell.Value)
    Next cell
End Sub


Sub MakeProperCase()
    Dim WorkRange As Range
    Dim cell As Range
'   Process only text cells, no formulas
    On Error Resume Next
    If Selection.Count = 1 Then
        Set WorkRange = ActiveCell
    Else
        Set WorkRange = Selection.SpecialCells _
        (xlCellTypeConstants, xlCellTypeConstants)
    End If
'   Proper case
    Application.ScreenUpdating = False
    For Each cell In WorkRange
        If Not cell.HasFormula Then cell.Value = WorksheetFunction.Proper(cell.Value)
    Next cell
End Sub
