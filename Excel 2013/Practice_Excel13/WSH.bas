Attribute VB_Name = "WSH"
Sub FileInfo()
    Dim fs As Object
    Dim objFile As Object
    Dim strMsg As String

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set objFile = fs.GetFile("C:\WINDOWS\System.ini")
    strMsg = "File name: " & _
        objFile.Name & vbCrLf
    strMsg = strMsg & "Disk: " & _
        objFile.Drive & vbCrLf
    strMsg = strMsg & "Date Created: " & _
        objFile.DateCreated & vbCrLf
    strMsg = strMsg & "Date Modified: " & _
        objFile.DateLastModified & vbCrLf
    MsgBox strMsg, , "File Information"
End Sub


Sub FileExists()
    Dim objFs As Object
    Dim strFile As String
    Set objFs = CreateObject("Scripting.FileSystemObject")
    strFile = InputBox("Enter the full name of the file: ")
    If objFs.FileExists(strFile) Then
        MsgBox strFile & " was found."
    Else
        MsgBox "File does not exist."
    End If
End Sub


Sub CopyFile()
    Dim objFs As Object
    Dim strFile As String
    Dim strNewFile As String

    strFile = "C:\Hello.doc"
    strNewFile = "C:\Program Files\Hello.doc"

    Set objFs = CreateObject("Scripting.FileSystemObject")
    objFs.CopyFile strFile, strNewFile
    MsgBox "A copy of the specified file was created."
    Set objFs = Nothing
End Sub

Sub DeleteFile()
    ' This procedure requires that you set up
    ' a reference to Microsoft Scripting Runtime
    ' Object Library by choosing Tools | References
    ' in the VBE window
    Dim objFs As FileSystemObject
    Set objFs = New FileSystemObject

    objFs.DeleteFile "C:\Program Files\Hello.doc"
    MsgBox "The requested file was deleted."
End Sub

Function DriveExists(disk)
    Dim objFs As Object
    Dim strMsg As String
    Set objFs = CreateObject("Scripting.FileSystemObject")
    If objFs.DriveExists(disk) Then
        strMsg = "Drive " & UCase(disk) & " exists."
    Else
        strMsg = UCase(disk) & " was not found."
    End If
    DriveExists = strMsg
' run this function from the worksheet
' by entering the following in any cell : =DriveExists("E:\")
End Function


Sub DriveInfo()
    Dim objFs As Object
    Dim objDisk As Object
    Dim infoStr As String
    Dim strDiskName As String
    strDiskName = InputBox("Enter the drive letter:", _
        "Drive Name", "C:\")

    Set objFs = CreateObject("Scripting.FileSystemObject")
    Set objDisk = objFs.GetDrive(objFs.GetDriveName(strDiskName))
    infoStr = "Drive: " & UCase(strDiskName) & vbCrLf
    infoStr = infoStr & "Drive letter: " & _
        UCase(objDisk.DriveLetter) & vbCrLf
    infoStr = infoStr & "Drive Type: " & objDisk.DriveType & vbCrLf
    infoStr = infoStr & "Drive File System: " & _
        objDisk.FileSystem & vbCrLf
    infoStr = infoStr & "Drive SerialNumber: " & _
        objDisk.SerialNumber & vbCrLf
    infoStr = infoStr & "Total Size in Bytes: " & _
        FormatNumber(objDisk.TotalSize / 1024, 0) & " Kb" & vbCrLf
    infoStr = infoStr & "Free Space on Drive: " & _
        FormatNumber(objDisk.FreeSpace / 1024, 0) & " Kb" & vbCrLf
    MsgBox infoStr, vbInformation, "Drive Information"
End Sub


Function DriveName(disk) As String
    Dim objFs As Object
    Dim strDiskName As String

    Set objFs = CreateObject("Scripting.FileSystemObject")
    strDiskName = objFs.GetDriveName(disk)
    DriveName = strDiskName
' run this function from the Immediate window
' by entering ?DriveName("C:\")
End Function

Sub DoesFolderExist()
    Dim objFs As Object
    Set objFs = CreateObject("Scripting.FileSystemObject")
    MsgBox objFs.FolderExists("C:\Program Files")
End Sub



Sub FilesInFolder()
    Dim objFs As Object
    Dim objFolder As Object
    Dim objFile As Object

    Set objFs = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFs.GetFolder("C:\")

    Workbooks.Add
    For Each objFile In objFolder.Files
        With ActiveCell
            .Formula = objFile.Name
            .Offset(0, 1).Range("A1").Formula = objFile.Type
            .Offset(1, 0).Range("A1").Select
        End With
    Next
    Columns("A:B").AutoFit
End Sub


Sub SpecialFolders()
    Dim objFs As Object
    Dim strWindowsFolder As String
    Dim strSystemFolder As String
    Dim strTempFolder As String

    Set objFs = CreateObject("Scripting.FileSystemObject")
    strWindowsFolder = objFs.GetSpecialFolder(0)
    strSystemFolder = objFs.GetSpecialFolder(1)
    strTempFolder = objFs.GetSpecialFolder(2)

    MsgBox strWindowsFolder & vbCrLf _
        & strSystemFolder & vbCrLf _
        & strTempFolder, vbInformation + vbOKOnly, _
        "Special Folders"
End Sub

Sub MakeNewFolder()
    Dim objFs As Object
    Dim objFolder As Object
    Set objFs = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFs.CreateFolder("C:\TestFolder")
    MsgBox "A new folder named " & _
        objFolder.Name & " was created."
End Sub


Sub MakeFolderCopy()
    Dim objFs As FileSystemObject
    Set objFs = New FileSystemObject
    If objFs.FolderExists("C:\TestFolder") Then
        objFs.CopyFolder "C:\TestFolder", "C:\FinalFolder"
        MsgBox "The folder was copied."
    End If
End Sub



Sub RemoveFolder()
    Dim objFs As Object
    Dim objFolder As Object
    Set objFs = CreateObject("Scripting.FileSystemObject")

    If objFs.FolderExists("C:\TestFolder") Then
        objFs.DeleteFolder "C:\TestFolder"
        MsgBox "The folder was deleted."
    End If
End Sub


Sub ReadTextFile()
    Dim objFs As Object
    Dim objFile As Object
    Dim strContent As String
    Dim strFileName As String

    strFileName = "C:\Windows\System.ini"
    Set objFs = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFs.OpenTextFile(strFileName)
    Do While Not objFile.AtEndOfStream
        strContent = strContent & objFile.ReadLine & vbCrLf
    Loop

    objFile.Close
    Set objFile = Nothing
    ActiveWorkbook.Sheets(3).Select
    Range("A1").Formula = strContent
    Columns("A:A").Select
    With Selection
        .ColumnWidth = 62.43
        .Rows.AutoFit
    End With
End Sub


Sub DrivesList()
    Dim objFs As Object
    Dim colDrives As Object
    Dim strDrive As String
    Dim Drive As Variant

    Set objFs = CreateObject("Scripting.FileSystemObject")
    Set colDrives = objFs.Drives

    For Each Drive In colDrives
        strDrive = "Drive " & Drive.DriveLetter & ": "
        Debug.Print strDrive
    Next
End Sub


Sub CountFilesInFolder()
    Dim objFs As Object
    Dim strFolder As String
    Dim objFolder As Object
    Dim objFiles As Object
    
    strFolder = InputBox("Enter the folder name:")
    If Not IsFolderEmpty(strFolder) Then
        Set objFs = CreateObject("Scripting.FileSystemObject")
        Set objFolder = objFs.GetFolder(strFolder)
        Set objFiles = objFolder.Files
        MsgBox "The number of files in the folder " & _
            strFolder & "=" & objFiles.Count
    Else
        MsgBox "Folder " & strFolder & " has 0 files."
    End If
End Sub

Function IsFolderEmpty(myFolder)
    Dim objFs As Object
    Dim objFolder As Object
    
    Set objFs = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFs.GetFolder(myFolder)
    IsFolderEmpty = (objFolder.Size = 0)
End Function



Sub CDROM_DriveLetter()
    Dim objFs As Object
    Dim colDrives As Object
    Dim Drive As Object
    Dim counter As Integer
    Const CDROM = 4
    
    Set objFs = CreateObject("Scripting.FileSystemObject")
    Set colDrives = objFs.Drives
    counter = 0
    For Each Drive In colDrives
        If Drive.DriveType = CDROM Then
            counter = counter + 1
            Debug.Print "The CD-ROM Drive: " & Drive.DriveLetter
        End If
    Next
    MsgBox "There are " & counter & " CD-ROM drives."
End Sub


Function IsCDROMReady(strDriveLetter)
    Dim objFs As Object
    Dim objDrive As Object

    Set objFs = CreateObject("Scripting.FileSystemObject")
    Set objDrive = objFs.GetDrive(strDriveLetter)

    IsCDROMReady = (objDrive.DriveType = 4) And _
        objDrive.IsReady = True
    ' run this function from the Immediate window
    ' by entering: ?IsCDROMReady("D:")
End Function

Sub CreateFile_Method1()
    Dim objFs As Object
    Dim objFile As Object
    
    Set objFs = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFs.CreateTextFile("C:\Phones.txt", True)

    objFile.WriteLine ("Margaret Kubiak: 212-338-8778")
    objFile.WriteBlankLines (2)
    objFile.WriteLine ("Robert Prochot: 202-988-2331")
    objFile.Close
End Sub


Sub CreateFile_Method2()
    Dim objFs As Object
    Dim objFile As Object
    
    Const ForWriting = 2
        
    Set objFs = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFs.OpenTextFile("C:\Shopping.txt", _
        ForWriting, True)
    
    objFile.WriteLine ("Bread")
    objFile.WriteLine ("Milk")
    objFile.WriteLine ("Strawberries")
    objFile.Close
End Sub


Sub CreateFile_Method3()
    Dim objFs As Object
    Dim objFile As Object
    Dim objText As Object
    Const ForWriting = 2
    Const ForReading = 1
    
    Set objFs = CreateObject("Scripting.FileSystemObject")
    objFs.CreateTextFile "New.txt"
    Set objFile = objFs.GetFile("New.txt")
    Set objText = objFile.OpenAsTextStream(ForWriting, _
        TristateUseDefault)
    
    objText.Write "Wedding Invitation"
    objText.Close
    Set objText = objFile.OpenAsTextStream(ForReading, _
        TristateUseDefault)
    MsgBox objText.ReadLine
    objText.Close
End Sub



