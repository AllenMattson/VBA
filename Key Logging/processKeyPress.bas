Attribute VB_Name = "processKeyPress"
Sub processKeyPress(keyPressed As String)

    'Select Case Asc(keyPressed)
    Select Case keyPressed
        'case 65 to 90, 97 to 122 'for Asc(keyPressed)
        Case "a" To "z", "A" To "Z":
            Application.StatusBar = "key Pressed was: " & keyPressed
            'put code here if capital alpha key was pressed
        Case "*":
            Call Unhook_KeyBoard
            Application.StatusBar = False
        Case Else
            'do whatever
    End Select
End Sub
