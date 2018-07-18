Attribute VB_Name = "Module1"
Option Explicit

Sub ZipFiles()
    Dim ShellApp As Object
    Dim FileNameZip As Variant
    Dim FileNames As Variant
    Dim i As Long, FileCount As Long

'   Get the file names
    FileNames = Application.GetOpenFilename _
        (FileFilter:="All Files (*.*),*.*", _
         FilterIndex:=1, _
         Title:="Select the files to ZIP", _
         MultiSelect:=True)

'   Exit if dialog box canceled
    If Not IsArray(FileNames) Then Exit Sub
   
    FileCount = UBound(FileNames)
    FileNameZip = Application.DefaultFilePath & "\compressed.zip"
    
    'Create empty Zip File with zip header
    Open FileNameZip For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
 
    Set ShellApp = CreateObject("Shell.Application")
    'Copy the files to the compressed folder
    For i = LBound(FileNames) To UBound(FileNames)
        DoEvents
        ShellApp.Namespace(FileNameZip).CopyHere FileNames(i)

        'Keep script waiting until Compressing is done
        On Error Resume Next
        Do Until ShellApp.Namespace(FileNameZip).items.Count = i
          DoEvents
          Application.Wait (Now + TimeValue("0:00:01"))
        Loop
        Application.StatusBar = "File " & i & " of " & UBound(FileNames)
    Next i
    
    Application.StatusBar = False
'   Prompt to view
    If MsgBox(FileCount & " files were zipped to:" & _
       vbNewLine & FileNameZip & vbNewLine & vbNewLine & _
       "Do you want to view the zip file?", vbQuestion + vbYesNo) = vbYes Then _
       Shell "Explorer.exe /e," & FileNameZip, vbNormalFocus
End Sub

