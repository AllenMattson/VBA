Attribute VB_Name = "WSH_Additional"
Sub RunNotepad()
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.Run "Notepad"
    Set WshShell = Nothing
End Sub


Sub OpenTxtFileInNotepad()
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    'wshShell.Run "Notepad C:\Phones.txt"
    'wshShell.Run "Control.exe"
    WshShell.Run "Control.exe Sysdm.cpl, ,2"
    Set WshShell = Nothing
End Sub

Sub ReadEnvVar()
    Dim WshShell As Object
    Dim objEnv As Object
    
    Set WshShell = CreateObject("WScript.Shell")
    Set objEnv = WshShell.Environment("Process")
    
    Debug.Print "Path=" & objEnv("PATH")
    Debug.Print "System Drive=" & objEnv("SYSTEMDRIVE")
    Debug.Print "System Root=" & objEnv("SYSTEMROOT")
    Debug.Print "Windows folder=" & objEnv("Windir")
    Debug.Print "Operating System=" & objEnv("OS")
    Set WshShell = Nothing
End Sub

Sub GetUserDomainComputer()
    Dim WshNetwork As Object
    Dim myData As String
    
    Set WshNetwork = CreateObject("WScript.Network")
    myData = myData & "Computer Name: " _
        & WshNetwork.ComputerName & vbCrLf
    myData = myData & "Domain: " _
        & WshNetwork.UserDomain & vbCrLf
    myData = myData & "User Name: " _
        & WshNetwork.UserName & vbCrLf
    
    MsgBox myData
End Sub


Sub CreateShortcut()
    ' this script creates two desktop shortcuts
    Dim WshShell As Object
    Dim objShortcut As Object
    Dim strWebAddr As String
    
    strWebAddr = "http://www.merclearning.com"

    Set WshShell = CreateObject("WScript.Shell")
    
    ' create an Internet shortcut
    Set objShortcut = WshShell.CreateShortcut(WshShell. _
        SpecialFolders("Desktop") & "\Mercury Learning.url")
    With objShortcut
        .TargetPath = strWebAddr
        .Save
    End With
    
    ' create a file shortcut
    ' you cannot create a shortcut to unsaved workbook file
    Set objShortcut = WshShell.CreateShortcut(WshShell. _
        SpecialFolders("Desktop") & "\" & ActiveWorkbook.Name & ".lnk")
    With objShortcut
        .TargetPath = ActiveWorkbook.FullName
        .Description = "Discover Mercury Learning"
        .WindowStyle = 7
        .Save
    End With

    Set objShortcut = Nothing
    Set WshShell = Nothing
End Sub


Sub ListShortcuts()
    Dim objFs As Object
    Dim objFolder As Object
    Dim WshShell As Object
    Dim strLinks As String
    Dim s As Variant
    Dim f As Variant
    
    Set WshShell = CreateObject("WScript.Shell")
    Set objFs = CreateObject("Scripting.FileSystemObject")
    strLinks = ""
    
    For Each s In WshShell.SpecialFolders
        Set objFolder = objFs.GetFolder(s)
        strLinks = strLinks & objFolder.Name _
            & " Shortcuts:" & vbCrLf

        If objFolder.Name = "Desktop" Then
            For Each f In objFolder.Files
                If InStrRev(UCase(f), ".LNK") Then
                    strLinks = strLinks & f.Name & vbCrLf
                End If
            Next
        End If
        Exit For
    Next
    Debug.Print strLinks
End Sub




