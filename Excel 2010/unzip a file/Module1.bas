Attribute VB_Name = "Module1"
Option Explicit

Sub UnzipAFile()
    Dim ShellApp As Object
    Dim TargetFile
    Dim ZipFolder

'   Target file & temp dir
    TargetFile = Application.GetOpenFilename _
        (FileFilter:="Zip Files (*.zip), *.zip")
    If TargetFile = False Then Exit Sub
    
    ZipFolder = Application.DefaultFilePath & "\Unzipped\"

'   Create a temp folder
    On Error Resume Next
    RmDir ZipFolder
    MkDir ZipFolder
    On Error GoTo 0

'   Copy the zipped files to the newly created folder
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(ZipFolder).CopyHere _
       ShellApp.Namespace(TargetFile).items

    If MsgBox("The files was unzipped to:" & _
       vbNewLine & ZipFolder & vbNewLine & vbNewLine & _
       "View the folder?", vbQuestion + vbYesNo) = vbYes Then _
       Shell "Explorer.exe /e," & ZipFolder, vbNormalFocus
End Sub


