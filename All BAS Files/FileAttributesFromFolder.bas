Attribute VB_Name = "FileAttributesFromFolder"
Option Explicit
Public X()
Public i As Long
Public objShell, objFolder, objFolderItem
Public FSO, oFolder, Fil
Sub FileAttributesFromFolder()
    Dim CSVWB As Workbook: Set CSVWB = Workbooks.Add
    Dim NewSht As Worksheet
    Dim MainFolderName As String
    Dim TimeLimit As Long, StartTime As Double
    Dim BottomRow As Long
    ReDim X(1 To 65535, 1 To 11)
     
    Set objShell = CreateObject("Shell.Application")
    TimeLimit = Application.InputBox("Please enter the maximum time that you wish this code to run for in minutes" & vbNewLine & vbNewLine & _
    "Leave this at zero for unlimited runtime", "Time Check box", 0)
    StartTime = Timer
     
    Application.ScreenUpdating = False
    MainFolderName = BrowseForFolder()
    Set NewSht = ThisWorkbook.Sheets.Add
     
    X(1, 1) = "Path"
    X(1, 2) = "File Name"
    X(1, 3) = "Last Accessed"
    X(1, 4) = "Last Modified"
    X(1, 5) = "Created"
    X(1, 6) = "Type"
    X(1, 7) = "Size"
    X(1, 8) = "Owner"
    X(1, 9) = "Author"
    X(1, 10) = "Title"
    X(1, 11) = "Comments"
     
    i = 1
     
    Set FSO = CreateObject("scripting.FileSystemObject")
    Set oFolder = FSO.GetFolder(MainFolderName)
     'error handling to stop the obscure error that occurs at time when retrieving DateLastAccessed
    On Error Resume Next
    For Each Fil In oFolder.Files
        Set objFolder = objShell.Namespace(oFolder.Path)
        Set objFolderItem = objFolder.ParseName(Fil.Name)
        i = i + 1
        If i Mod 20 = 0 And TimeLimit <> 0 And Timer > (TimeLimit * 60 + StartTime) Then
            GoTo FastExit
        End If
        If i Mod 50 = 0 Then
            Application.StatusBar = "Processing File " & i
            DoEvents
        End If
        X(i, 1) = oFolder.Path
        X(i, 2) = Fil.Name
        X(i, 3) = Fil.DateLastAccessed
        X(i, 4) = Fil.DateLastModified
        X(i, 5) = Fil.DateCreated
        X(i, 6) = Fil.Type
        X(i, 7) = Fil.Size
        X(i, 8) = objFolder.GetDetailsOf(objFolderItem, 8)
        X(i, 9) = objFolder.GetDetailsOf(objFolderItem, 9)
        X(i, 10) = objFolder.GetDetailsOf(objFolderItem, 10)
        X(i, 11) = objFolder.GetDetailsOf(objFolderItem, 14)
    Next
     
     'Get subdirectories
    If TimeLimit = 0 Then
        Call RecursiveFolder(oFolder, 0)
    Else
        If Timer < (TimeLimit * 60 + StartTime) Then
        Call RecursiveFolder(oFolder, TimeLimit * 60 + StartTime)
        End If
    End If
    

FastExit:
With CSVWB
    Range("A:K") = X
    If i < 65535 Then Range(Cells(i + 1, "A"), Cells(65536, "A")).EntireRow.Delete
        Cells.Replace What:="#N/A", Replacement:="", LookAt:=xlPart, SearchOrder _
        :=xlByColumns, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Range("A:K").WrapText = False
    Range("A:K").EntireColumn.AutoFit
    Range("1:1").Font.Bold = True
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    Range("a1").Activate
End With
    Set FSO = Nothing
    Set objShell = Nothing
    Set oFolder = Nothing
    Set objFolder = Nothing
    Set objFolderItem = Nothing
    Set Fil = Nothing
    Set WB = Nothing
    Set CSVWB = Nothing
    Application.StatusBar = ""
    Application.ScreenUpdating = True
End Sub
 
Sub RecursiveFolder(xFolder, TimeTest As Long)
    Dim SubFld
    For Each SubFld In xFolder.SubFolders
        Set oFolder = FSO.GetFolder(SubFld)
        Set objFolder = objShell.Namespace(SubFld.Path)
        For Each Fil In SubFld.Files
            Set objFolder = objShell.Namespace(oFolder.Path)
             'Problem with objFolder at times
            If Not objFolder Is Nothing Then
                Set objFolderItem = objFolder.ParseName(Fil.Name)
                i = i + 1
                Debug.Print TheBigCount & " : " & Fil.Name
                If i Mod 20 = 0 And TimeTest <> 0 And Timer > TimeTest Then
                    Exit Sub
                End If
                If i Mod 50 = 0 Then
                    Application.StatusBar = "Processing File " & i
                    DoEvents
                End If
                X(i, 1) = SubFld.Path
                X(i, 2) = Fil.Name
                X(i, 3) = Fil.DateLastAccessed
                X(i, 4) = Fil.DateLastModified
                X(i, 5) = Fil.DateCreated
                X(i, 6) = Fil.Type
                X(i, 7) = Fil.Size
                X(i, 8) = objFolder.GetDetailsOf(objFolderItem, 8)
                X(i, 9) = objFolder.GetDetailsOf(objFolderItem, 9)
                X(i, 10) = objFolder.GetDetailsOf(objFolderItem, 10)
                X(i, 11) = objFolder.GetDetailsOf(objFolderItem, 14)
            Else
                'Debug.Print Fil.Path & " " & Fil.Name
            End If
        Next
        Call RecursiveFolder(SubFld, TimeTest)
    Next
End Sub
Function BrowseForFolder(Optional OpenAt As Variant) As Variant
     'Function purpose:  To Browser for a user selected folder.
     'If the "OpenAt" path is provided, open the browser at that directory
     'NOTE:  If invalid, it will open at the Desktop level
     
    Dim ShellApp As Object
     
     'Create a file browser window at the default folder
    Set ShellApp = CreateObject("Shell.Application"). _
    BrowseForFolder(0, "Please choose a folder", 0, OpenAt)
     
     'Set the folder to that selected.  (On error in case cancelled)
    On Error Resume Next
    BrowseForFolder = ShellApp.self.Path
    On Error GoTo 0
     
     'Destroy the Shell Application
    Set ShellApp = Nothing
     
     'Check for invalid or non-entries and send to the Invalid error
     'handler if found
     'Valid selections can begin L: (where L is a letter) or
     '\\ (as in \\servername\sharename.  All others are invalid
    Select Case Mid(BrowseForFolder, 2, 1)
    Case Is = ":"
        If Left(BrowseForFolder, 1) = ":" Then GoTo Invalid
    Case Is = "\"
        If Not Left(BrowseForFolder, 1) = "\" Then GoTo Invalid
    Case Else
        GoTo Invalid
    End Select
     
    Exit Function
     
Invalid:
     'If it was determined that the selection was invalid, set to False
    BrowseForFolder = False
     
End Function



