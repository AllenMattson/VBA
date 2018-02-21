Attribute VB_Name = "Traps"
Sub Archive()
    Dim folderName As String
    Dim MyDrive As String
    Dim BackupName As String
    Application.DisplayAlerts = False
    
    On Error GoTo ErrorHandler
    
    folderName = ActiveWorkbook.Path
    
    If folderName = "" Then
        MsgBox "You can't copy this file. " & Chr(13) _
            & "This file has not been saved.", _
        vbInformation, "File Archive"
    Else
        With ActiveWorkbook
            If Not .Saved Then .Save
            MyDrive = InputBox("Enter the Pathname:" & _
                Chr(13) & "(for example: D:\, " & _
                    "E:\MyFolder\, etc.)", _
                    "Archive Location?", "D:\")
            If MyDrive <> "" Then
                If Right(MyDrive, 1) <> "\" Then
                    MyDrive = MyDrive & "\"
                End If
                BackupName = MyDrive & .Name
                .SaveCopyAs Filename:=BackupName
                MsgBox .Name & " was copied to: " _
                    & MyDrive, , "End of Archiving"
            End If
        End With
  End If
  GoTo ProcEnd
ErrorHandler:
    MsgBox "Visual Basic cannot find the " & _
        "specified path (" & MyDrive & ")" & Chr(13) & _
        "for the archive. Please try again.", _
        vbInformation + vbOKOnly, "Disk Drive or " & _
        "Folder does not exist"
ProcEnd:
    Application.DisplayAlerts = True
End Sub


Sub OpenToRead()
Dim myFile As String
Dim myChar As String
Dim myText As String
Dim FileExists As Boolean

    FileExists = True

    On Error GoTo ErrorHandler

    myFile = InputBox("Enter the name of file you want to open:")
    Open myFile For Input As #1
    If FileExists Then
        Do While Not EOF(1)         ' loop until the end of file
            myChar = Input(1, #1)   ' get one character
            myText = myText + myChar ' store in the variable myText
        Loop
        Debug.Print myText      ' print to the Immediate window
        Close #1                ' close the file
    End If
    Exit Sub

ErrorHandler:
    FileExists = False
    Select Case Err.Number
        Case 76
        MsgBox "The path you entered cannot be found."
        Case 53
        MsgBox "This file can't be found on the " & _
            "specified drive."
        Case 75
            Exit Sub
        Case Else
            MsgBox "Error " & Err.Number & " :" & Error(Err.Number)
        Exit Sub
    End Select
    Resume Next
End Sub


