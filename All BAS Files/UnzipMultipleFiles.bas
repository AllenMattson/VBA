Attribute VB_Name = "UnzipMultipleFiles"
Option Explicit

Sub UnzipMultipleFiles()
Application.DisplayAlerts = False

    Dim ShellApp As Object
    Dim TargetFile
    Dim NewTargetFile
    Dim ZipFolder
    Dim i As Integer
'   Target file & temp dir
    TargetFile = Application.GetOpenFilename _
        (FileFilter:="Zip Files (*.zip), *.zip", Title:="Select Zip Files", MultiSelect:=True)
     ZipFolder = Application.DefaultFilePath & "\Unzipped\"

'   Create a temp folder
    On Error Resume Next
    RmDir ZipFolder
    MkDir ZipFolder
    On Error GoTo 0
    For i = 1 To UBound(TargetFile)
        NewTargetFile = TargetFile(i)
'    If TargetFile = False Then Exit Sub
    


'   Copy the zipped files to the newly created folder
    Set ShellApp = CreateObject("Shell.Application")
    ShellApp.Namespace(ZipFolder).CopyHere _
       ShellApp.Namespace(NewTargetFile).items

    Next i
    
    If MsgBox("The files was unzipped to:" & _
       vbNewLine & ZipFolder & vbNewLine & vbNewLine & _
       "View the folder?", vbQuestion + vbYesNo) = vbYes Then _
       Shell "Explorer.exe /e," & ZipFolder, vbNormalFocus
Application.DisplayAlerts = True
End Sub

