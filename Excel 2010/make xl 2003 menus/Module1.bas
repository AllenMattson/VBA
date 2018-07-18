Attribute VB_Name = "Module1"
Sub MakeOldMenus()
    Dim cb As CommandBar
    Dim cbc As CommandBarControl
    Dim OldMenu As CommandBar

'   Delete it, if it exists
    On Error Resume Next
    Application.CommandBars("Old Menus").Delete
    On Error GoTo 0

'   Create an old-style toolbar
'   Set the last argument to False for a more compact menu
    Set OldMenu = Application.CommandBars.Add("Old Menus", , True)
    
'   Copy the controls from Excel's "Built-in Menus" shortcut menu
    With CommandBars("Built-in Menus")
        .Controls("&File").Copy OldMenu
        .Controls("&Edit").Copy OldMenu
        .Controls("&View").Copy OldMenu
        .Controls("&Insert").Copy OldMenu
        .Controls("F&ormat").Copy OldMenu
        .Controls("&Tools").Copy OldMenu
        .Controls("&Data").Copy OldMenu
        .Controls("&Window").Copy OldMenu
        .Controls("&Help").Copy OldMenu
    End With

'   Make it visible. It appears in the Add-Ins tab
    Application.CommandBars("Old Menus").Visible = True
End Sub

