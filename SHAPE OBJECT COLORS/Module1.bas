Attribute VB_Name = "Module1"
Option Explicit
Sub AddToShortCut()
'   Adds a menu item to the Cell shortcut menu (active workbook)
    Dim Bar As CommandBar
    Dim NewControl As CommandBarButton
    DeleteFromShortcut
    Set Bar = CommandBars("Cell")
    Set NewControl = Bar.Controls.Add _
        (Type:=msoControlButton)
    With NewControl
        .Caption = "Toggle &Wrap Text"
        .OnAction = "ToggleWrapText"
        .Picture = Application.CommandBars.GetImageMso _
           ("WrapText", 16, 16)
        .Style = msoButtonIconAndCaption
    End With
End Sub

Sub DeleteFromShortcut()
    On Error Resume Next
    CommandBars("Cell").Controls("Toggle &Wrap Text").Delete
End Sub




Sub AddToShortCut2()
'   Adds a menu item to the Cell shortcut menu (all open workbook windows)
    Dim NewControl As CommandBarButton
    Dim activeWin As Window
    Dim w As Window
    
    Set activeWin = ActiveWindow
    Application.ScreenUpdating = False
   
'   Loop through each visible window
    For Each w In Windows
        If w.Visible Then
            w.Activate
'           Delete control if it already exists
            On Error Resume Next
            CommandBars("Cell").Controls("Toggle &Wrap Text").Delete
            On Error GoTo 0
'           Add the new control
            Set NewControl = CommandBars("Cell").Controls.Add _
                (Type:=msoControlButton, ID:=1)
            With NewControl
                .Caption = "Toggle &Wrap Text"
                .OnAction = "ToggleWrapText"
                .Picture = Application.CommandBars.GetImageMso("WrapText", 16, 16)
                .Style = msoButtonIconAndCaption
            End With
        End If
    Next w
'   Activate original window
    activeWin.Activate
    Application.ScreenUpdating = True
End Sub


Sub DeleteFromShortcut2()
'   All open workbook windows
    Dim activeW As Window
    Dim w As Window
    Set activeW = ActiveWindow
    On Error Resume Next
    Application.ScreenUpdating = False
    For Each w In Windows
        If w.Visible Then
            w.Activate
            CommandBars("Cell").Controls("Toggle &Wrap Text").Delete
        End If
    Next w
    activeW.Activate
End Sub


Sub ToggleWrapText()
    On Error Resume Next
    CommandBars.ExecuteMso "WrapText"
    If Err.Number <> 0 Then MsgBox "Could not toggle Wrap Text"
End Sub

