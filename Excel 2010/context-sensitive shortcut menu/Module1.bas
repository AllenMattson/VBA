Attribute VB_Name = "Module1"
Option Explicit

Sub CreateShortcut()
    Dim myBar As CommandBar
    Dim myItem As CommandBarControl
    
    DeleteShortcut
    Set myBar = CommandBars.Add _
      (Name:="MyShortcut", Position:=msoBarPopup, Temporary:=True)
    
    Set myItem = myBar.Controls.Add(Type:=msoControlButton)
    With myItem
        .Caption = "&Number Format..."
        .OnAction = "ShowFormatNumber"
        .FaceId = 1554
    End With
        
    Set myItem = myBar.Controls.Add(Type:=msoControlButton)
    With myItem
        .Caption = "&Alignment..."
        .OnAction = "ShowFormatAlignment"
        .FaceId = 194
    End With
        
    Set myItem = myBar.Controls.Add(Type:=msoControlButton)
    With myItem
        .Caption = "&Font..."
        .OnAction = "ShowFormatFont"
        .FaceId = 309
    End With

    Set myItem = myBar.Controls.Add(Type:=msoControlButton)
    With myItem
        .Caption = "&Borders..."
        .OnAction = "ShowFormatBorder"
        .FaceId = 149
        .BeginGroup = True
    End With
    
    Set myItem = myBar.Controls.Add(Type:=msoControlButton)
    With myItem
        .Caption = "&Fill..."
        .OnAction = "ShowFormatPatterns"
        .FaceId = 687
    End With
    
    Set myItem = myBar.Controls.Add(Type:=msoControlButton)
    With myItem
        .Caption = "&Protection..."
        .OnAction = "ShowFormatProtection"
        .FaceId = 225
    End With
End Sub


Sub ShowFormatNumber()
    'Application.Dialogs(xlDialogFormatNumber).Show
    CommandBars.ExecuteMso ("FormatCellsNumberDialog")
End Sub

Sub ShowFormatAlignment()
    'Application.Dialogs(xlDialogAlignment).Show
    CommandBars.ExecuteMso ("CellAlignmentOptions")
End Sub

Sub ShowFormatFont()
    'Application.Dialogs(xlDialogFormatFont).Show
    CommandBars.ExecuteMso ("FormatCellsFontDialog")
End Sub

Sub ShowFormatBorder()
    'Application.Dialogs(xlDialogBorder).Show
    CommandBars.ExecuteMso ("BordersMOreDialog")
End Sub

Sub ShowFormatPatterns()
    '(No ExecuteMso command, so use old-style)
    Application.Dialogs(xlDialogPatterns).Show
End Sub

Sub ShowFormatProtection()
    '(No ExecuteMso command, so use old-style)
    Application.Dialogs(xlDialogCellProtection).Show
End Sub

Sub DeleteShortcut()
    On Error Resume Next
    CommandBars("MyShortcut").Delete
End Sub

Sub ShowMyShortcutMenu()
Attribute ShowMyShortcutMenu.VB_ProcData.VB_Invoke_Func = "M\n14"
'   Ctrl+Shift+M shortcut key
    CommandBars("MyShortcut").ShowPopup
End Sub

