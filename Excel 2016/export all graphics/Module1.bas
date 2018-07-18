Attribute VB_Name = "Module1"
Option Explicit

Sub SaveAllGraphics()
    Dim FileName As String
    Dim TempName As String
    Dim DirName As String
    Dim gFile As String
    
    FileName = ActiveWorkbook.FullName
    TempName = ActiveWorkbook.Path & "\" & ActiveWorkbook.Name & "graphics.htm"
    DirName = Left(TempName, Len(TempName) - 4) & "_files"
    
'   Save active workbookbook as HTML, then reopen original
    ActiveWorkbook.Save
    ActiveWorkbook.SaveAs FileName:=TempName, FileFormat:=xlHtml
    Application.DisplayAlerts = False
    ActiveWorkbook.Close
    Workbooks.Open FileName
    
'   Delete the HTML file
    Kill TempName
    
'   Delete all but *.PNG files in the HTML folder
    gFile = Dir(DirName & "\*.*")
    Do While gFile <> ""
        If Right(gFile, 3) <> "png" Then Kill DirName & "\" & gFile
        gFile = Dir
    Loop

'   Show the exported graphics
    Shell "explorer.exe " & DirName, vbNormalFocus
End Sub


