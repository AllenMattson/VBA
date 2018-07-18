Attribute VB_Name = "KillStatement"
Sub RemoveMe()
    Dim folder As String
    Dim myfile As String

    ' assign the name of folder to the folder variable
    ' notice the ending backslash "\"
    folder = "C:\Abort\"
    myfile = Dir(folder, vbNormal)

    Do While myfile <> ""
        Kill folder & myfile
        myfile = Dir
    Loop
    RmDir folder
End Sub

