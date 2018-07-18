Attribute VB_Name = "MenuMod"
Option Explicit

Sub MakeMenu()
Attribute MakeMenu.VB_ProcData.VB_Invoke_Func = " \n14"
'   Adds shortcut menu items
    Dim Cap(1 To 15)
    Dim Mac(1 To 15)
    Dim NewMenu As CommandBarControl
    Dim Item As CommandBarControl
    Dim MenuCount As Integer
    Dim i As Integer
    
    Cap(1) = "Select Down (As In Ctrl+Shift+Down)"
    Mac(1) = "SelectDown"
    Cap(2) = "Select Up (As In Ctrl+Shift+Up)"
    Mac(2) = "SelectUp"
    Cap(3) = "Select To Right (As In Ctrl+Shift+Right)"
    Mac(3) = " SelectToRight"
    Cap(4) = "Select To Left (As In Ctrl+Shift+Left)"
    Mac(4) = " SelectToLeft"
    Cap(5) = "Select Current Region (As In Ctrl+Shift+*)"
    Mac(5) = " SelectCurrentRegion"
    Cap(6) = "Select Active Area (As In End, Home, Ctrl+Shift+Home)"
    Mac(6) = " SelectActiveArea"
    Cap(7) = "Select Contiguous Cells in ActiveCell's Column"
    Mac(7) = " SelectActiveColumn"
    Cap(8) = "Select Contiguous Cells in ActiveCell's Row"
    Mac(8) = " SelectActiveRow"
    Cap(9) = "Select an Entire Column (As In Ctrl+Spacebar)"
    Mac(9) = " SelectEntireColumn"
    Cap(10) = "Select an Entire Row  (As In Shift+Spacebar)"
    Mac(10) = " SelectEntireRow"
    Cap(11) = "Select the Entire Worksheet (As In Ctrl+A)"
    Mac(11) = " SelectEntireSheet"
    Cap(12) = "Activate the Next Blank Cell Below"
    Mac(12) = " ActivateNextBlankDown"
    Cap(13) = "Activate the Next Blank Cell To the Right"
    Mac(13) = " ActivateNextBlankToRight"
    Cap(14) = "Select From the First NonBlank to the Last Nonblank in the Row"
    Mac(14) = " SelectFirstToLastInRow"
    Cap(15) = "Select From the First NonBlank to the Last Nonblank in the Column"
    Mac(15) = " SelectFirstToLastInColumn"


'   Delete the menu if it already exists
    On Error Resume Next
    Application.CommandBars("Cell").Controls("&Selection Demo").Delete
    On Error GoTo 0
    
'   Add the menu
    MenuCount = Application.CommandBars("Cell").Controls.Count
    Set NewMenu = Application.CommandBars("Cell").Controls.Add(Type:=msoControlPopup, Before:=MenuCount, temporary:=True)
    NewMenu.Caption = "&Selection Demo"
    
'   Add the menu items
    For i = 1 To 15
        Set Item = NewMenu.Controls.Add(Type:=msocontrolbutton, temporary:=True)
        With Item
            .Caption = Cap(i)
            .OnAction = Mac(i)
            If i Mod 4 = 0 Then .BeginGroup = True
        End With
    Next i
End Sub

Sub DeleteMenu()
Attribute DeleteMenu.VB_ProcData.VB_Invoke_Func = " \n14"
'   Delete the menu before closing
    On Error Resume Next
    Application.CommandBars(1).Controls("Selection Demo").Delete
End Sub

