Attribute VB_Name = "DirFunction"
Sub MyFiles()
    Dim myfile As String
    Dim mpath As String

    mpath = InputBox("Enter pathname, e.g. C:\Excel2013_ByExample")
    If Right(mpath, 1) <> "\" Then mpath = mpath & "\"

    myfile = Dir(mpath & "*.*")
    If myfile <> "" Then Debug.Print "Files in the " & _
        mpath & " folder:"
    Debug.Print LCase$(myfile)
    If myfile = "" Then
        MsgBox "No files found."
        Exit Sub
    End If
    Do While myfile <> ""
        myfile = Dir
        Debug.Print LCase$(myfile)
    Loop
End Sub


Sub GetFiles()
    Dim myfile As String
    Dim nextRow As Integer

    nextRow = 1
    With Worksheets("Sheet1").Range("A1")
        myfile = Dir("C:\Excel2013_ByExample\*.*", vbNormal)
        .Value = myfile
        Do While myfile <> ""
            myfile = Dir
            .Offset(nextRow, 0).Value = myfile
            nextRow = nextRow + 1
        Loop
    End With
End Sub


