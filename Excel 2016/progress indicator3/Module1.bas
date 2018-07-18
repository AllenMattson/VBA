Attribute VB_Name = "Module1"
Option Explicit

Sub ProcessFiles()
    Dim sFile As String, lFile As Long
    Const sPATH As String = "C:\Text Files\"
    
    sFile = Dir(sPATH & "*.txt")
    UProgress.Caption = "Proccesing File Progress"
    UProgress.SetDescription "Completed files..."
    UProgress.Show vbModeless
    
    If Len(sFile) = 0 Then MsgBox ("Create a few text files in C:\Text Files\")
    
    Do While Len(sFile) > 0
        ImportFile sFile
        UProgress.AddStep sPATH & sFile
        sFile = Dir
    Loop
    Unload UProgress
End Sub

Sub ImportFile(sFile As String)
    Dim i As Long
    For i = 1 To 5000
        DoEvents
    Next i
End Sub

