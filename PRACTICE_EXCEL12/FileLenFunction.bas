Attribute VB_Name = "FileLenFunction"
Sub TotalBytesIni()
    Dim iniFile As String
    Dim allBytes As Long

    iniFile = Dir("C:\WINDOWS\*.ini")
    allBytes = 0
    Do While iniFile <> ""
        allBytes = allBytes + FileLen("C:\WINDOWS\" & iniFile)
        iniFile = Dir
    Loop
    Debug.Print "Total bytes: " & allBytes
End Sub


