Attribute VB_Name = "basTotxt"
Sub GetCodeFiles()
'C:\Users\Allen\Documents\GitHub\VBA
Dim FileName As String
Dim strPath As String
Dim strPath2 As String: strPath2 = "C:\Users\Allen\Desktop\Bas Files\"
Dim TargetFile As String
Dim ws As Worksheet: Set ws = Sheets("Sheet3")
Dim LR As Long
Dim i As Long
With ws
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To LR
        strPath = Cells(i, 1).Value
        TargetFile = Cells(i, 1).Offset(0, 1).Value
        If InStr(TargetFile, "Module") Then
            FileName = Cells(i, 1).Offset(0, 2).Value
            FileName = Replace(FileName, " ", "_")
        Else
            FileName = Cells(i, 1).Offset(0, 1).Value
        End If
        If Right(FileName, 4) = ".bas" Then FileName = Replace(FileName, ".bas", "")
        FileCopy strPath & TargetFile, strPath2 & FileName & ".txt"
     Next i
End With
End Sub
Sub RemoveFirstLine(strFileName As String)
'C:\Users\Allen\Desktop
Const FOR_READING = 1
Const FOR_WRITING = 2
'strFileName = "C:\Users\Allen\Desktop\Bas Files\ABOUT_RANGE_SELECTION.txt"
iNumberOfLinesToDelete = 1
 
Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(strFileName, FOR_READING)
strContents = objTS.ReadAll
objTS.Close
 
arrLines = Split(strContents, vbNewLine)
Set objTS = objFS.OpenTextFile(strFileName, FOR_WRITING)
 
For i = 0 To UBound(arrLines)
   If i > (iNumberOfLinesToDelete - 1) Then
      objTS.WriteLine arrLines(i)
   End If
Next
End Sub
Sub LoopFiles()

Dim CurrDir As String: CurrDir = "C:\Users\Allen\Desktop\Bas Files\"
'   Get files
    On Error Resume Next
    FileName = Dir(CurrDir & "*.txt", vbDirectory)
    Do While Len(FileName) <> 0
      If Left(FileName, 1) <> "." Then 'Current dir
      If FileName = "ABOUT_RANGE_SELECTION.txt" Then GoTo AlreadyDone
        PathAndName = CurrDir & FileName
        RemoveFirstLine (PathAndName)
    End If
AlreadyDone:
        FileName = Dir()
    Loop
End Sub

