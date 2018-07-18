Attribute VB_Name = "FileCopyAndKill"
Sub CopyToAbortFolder()
    Dim folder As String
    Dim source As String
    Dim dest As String
    Dim msg1 As String
    Dim msg2 As String
    Dim p As Integer
    Dim s As Integer
    Dim i As Long

    On Error GoTo ErrorHandler

    folder = "C:\Abort"
    msg1 = "The selected file is already in this folder."
    msg2 = "was copied to"
    p = 1
    i = 1
    ' get the name of the file from the user
    source = Application.GetOpenFilename
    ' don't do anything if cancelled
    If source = "False" Then Exit Sub
    ' get the total number of backslash characters "\" in the source
    ' variable's contents
    Do Until p = 0
        p = InStr(i, source, "\", 1)
        If p = 0 Then Exit Do
        s = p
        i = p + 1
    Loop
    ' create the destination filename
    dest = folder & Mid(source, s, Len(source))
        ' create a new folder with this name
        MkDir folder
        ' check if the specified file already exists in the
        ' destination folder
        If Dir(dest) <> "" Then
            MsgBox msg1
        Else
        ' copy the selected file to the C:\Abort folder
            FileCopy source, dest
            MsgBox source & " " & msg2 & " " & dest
        End If
        Exit Sub
ErrorHandler:
        If Err = "75" Then
            Resume Next
        End If
        If Err = "70" Then
            MsgBox "You can't copy an open file."
        Exit Sub
    End If
End Sub

