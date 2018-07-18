Attribute VB_Name = "Module1"
Option Explicit

Sub ShowInstalledFonts()
    Dim FontList As CommandBarControl
    Dim TempBar As CommandBar
    Dim i As Long
    
'   Create temporary CommandBar
    Set TempBar = Application.CommandBars.Add
    Set FontList = TempBar.Controls.Add(ID:=1728)
    
'   Put the fonts into column A
    Range("A:A").ClearContents
    For i = 0 To FontList.ListCount - 1
        Cells(i + 1, 1) = FontList.List(i + 1)
'        Cells(i + 1, 1).Font.Name = FontList.List(i + 1)
    Next i
    
'   Delete temporary CommandBar
    TempBar.Delete
End Sub


