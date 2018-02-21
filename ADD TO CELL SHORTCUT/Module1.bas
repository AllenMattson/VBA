Attribute VB_Name = "Module1"
Option Explicit

Sub AddToShortCut()
'   Adds a menu item to the Cell shortcut menu
    Dim Bar As CommandBar
    Dim NewControl As CommandBarButton
    DeleteFromShortcut
    Set Bar = CommandBars("Cell")
    Set NewControl = Bar.Controls.Add _
        (Type:=msoControlButton, ID:=1, _
         temporary:=True)
    With NewControl
        .Caption = "Toggle &Word Wrap"
        .OnAction = "ToggleWordWrap"
        .Picture = Application.CommandBars.GetImageMso("WrapText", 16, 16)
        .Style = msoButtonIconAndCaption
    End With
End Sub

Sub ToggleWordWrap()
    CommandBars.ExecuteMso ("WrapText")
End Sub

Sub DeleteFromShortcut()
    On Error Resume Next
    CommandBars("Cell").Controls _
      ("Toggle &Word Wrap").Delete
End Sub


