Attribute VB_Name = "Module1"
' Declare a module-level variable
Public blnIsFileSelected As Boolean

Sub UnzipExcelFile()
    Dim objShell As Object
    Dim strZipFile, strZipFolder, strSourceFile, objFile
    Dim strStartDir As String

    strStartDir = "C:\Excel2013_ByExample"
    'change folder
    If ActiveWorkbook.Path <> strStartDir Then
        ChDir strStartDir
    End If

    ' get Excel file to unzip
    strSourceFile = Application.GetOpenFilename _
    (FileFilter:="Excel Files (*.xlsx; *.xlsm), *.xlsx; *.xlsm", _
    Title:="Select Excel file you want to unzip")

    'exit if file was not selected
    If strSourceFile = False Then
        blnIsFileSelected = False
        Exit Sub
    End If

    blnIsFileSelected = True
    strZipFile = strSourceFile & ".zip"

    'create the zip file
    FileCopy strSourceFile, strZipFile

    'Create new folder to store unzipped files
    strZipFolder = "C:\Excel2013_ByExample\ZipPackage"
    On Error Resume Next
    MkDir strZipFolder

    'Copy package files to the ZipPackage folder
    Set objShell = CreateObject("Shell.Application")

    For Each objFile In objShell.Namespace(strZipFile).items
        objShell.Namespace(strZipFolder).CopyHere (objFile)
    Next objFile

    'Activate Windows Explorer
    Shell "Explorer.exe /e," & strZipFolder, vbNormalFocus

    'remove the zip file and release resources
    Kill strZipFile
    Set objShell = Nothing
End Sub


Sub CreateEmptyZipFile(strFileName As String)
    Dim strHeader As String
    Dim fso As Object

    strHeader = Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)

    ' delete the file if it already exists
    If Len(Dir(strFileName)) > 0 Then
        Kill strFileName
    End If

    ' add a required header
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile(strFileName).Write strHeader

End Sub


Sub ZipToExcel()
    Dim objShell As Object
    Dim strZipFile, strZipFolder, objFile
    Dim strStartDir As String
    Dim strExcelFile As String
    Dim mFlag As Boolean

    strZipFolder = "C:\Excel2013_ByExample\ZipPackage"
    strZipFile = "C:\Excel2013_ByExample\PackageModified.zip"
    mFlag = False

    'check if folder is empty
    If Len(Dir(strZipFolder & "\*.*")) < 1 Then
        MsgBox "There are no files to zip."
        Exit Sub
    End If

    ' check if a VBA project exists
    If Len(Dir(strZipFolder & "\xl\vbaProject.bin")) > 0 Then
        mFlag = True
    End If

    'Create an empty zip file
    CreateEmptyZipFile (strZipFile)

    'Copy files from strZipFolder to the strZipFile
    On Error Resume Next
    Set objShell = CreateObject("Shell.Application")

    For Each objFile In objShell.Namespace(strZipFolder).items
        objShell.Namespace(strZipFile).CopyHere (objFile)
        Application.Wait (Now + TimeValue("0:00:10"))
    Next objFile

    'Create Excel file name
    If mFlag Then
        strExcelFile = Replace(strZipFile, ".zip", ".xlsm")
    Else
        strExcelFile = Replace(strZipFile, ".zip", ".xlsx")
    End If

    'Rename the strZipFile
    Name strZipFile As strExcelFile

    Set objShell = Nothing
    Set objFile = Nothing

    MsgBox "Zipping files completed."
End Sub

