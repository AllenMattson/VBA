Attribute VB_Name = "DownloadInternetFiles"
Option Explicit

'API function declaration for both 32 and 64bit Excel.
#If VBA7 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                                    (ByVal pCaller As Long, _
                                    ByVal szURL As String, _
                                    ByVal szFileName As String, _
                                    ByVal dwReserved As Long, _
                                    ByVal lpfnCB As Long) As Long
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" _
                            (ByVal pCaller As Long, _
                            ByVal szURL As String, _
                            ByVal szFileName As String, _
                            ByVal dwReserved As Long, _
                            ByVal lpfnCB As Long) As Long
#End If
 
Sub DownloadFiles()
                    
    '--------------------------------------------------------------------------------------------------
    'The macro loops through all the URLs (column C) and downloads the files at the specified folder.
    'The given file names (column D) are used to create the full path of the files.
    'If the file is downloaded successfully an OK will appear in column E (otherwise an ERROR value).
    'The code is based on API function URLDownloadToFile, which actually does all the work.
            
    'Written By:    Christos Samaras
    'Date:          28/05/2014
    'Last Update:   06/06/2015
    'E-mail:        xristos.samaras@gmail.com
    'Site:          http://www.myengineeringworld.net
    '--------------------------------------------------------------------------------------------------
    
    'Declaring the necessary variables.
    Dim sh                  As Worksheet
    Dim DownloadFolder      As String
    Dim LastRow             As Long
    Dim SpecialChar()       As String
    Dim SpecialCharFound    As Double
    Dim FilePath            As String
    Dim i                   As Long
    Dim j                   As Integer
    Dim Result              As Long
    Dim CountErrors         As Long
    
    'Disable screen flickering.
    Application.ScreenUpdating = False
    
    'Set the worksheet object to the desired sheet.
    Set sh = Sheets("Main")
    
    'An array with special characters that cannot be used for naming a file.
    SpecialChar() = Split("\ / : * ? " & Chr$(34) & " < > |", " ")
    
    'Find the last row.
     With sh
        .Activate
        LastRow = .Cells(.Rows.Count, "C").End(xlUp).Row
    End With
    
    'Check if the download folder exists.
    DownloadFolder = sh.Range("B4")
    On Error Resume Next
    If Dir(DownloadFolder, vbDirectory) = vbNullString Then
        MsgBox "The folder's path is incorrect!", vbCritical, "Folder's Path Error"
        sh.Range("B4").Select
        Exit Sub
    End If
    On Error GoTo 0
               
    'Check if there is at least one URL.
    If LastRow < 8 Then
        MsgBox "You did't enter a single URL!", vbCritical, "No URL Error"
        sh.Range("C8").Select
        Exit Sub
    End If
    
    'Clear the results column.
    sh.Range("E8:E" & LastRow).ClearContents
    
    'Add the backslash if doesn't exist.
    If Right(DownloadFolder, 1) <> "\" Then
        DownloadFolder = DownloadFolder & "\"
    End If

    'Counting the number of files that will not be downloaded.
    CountErrors = 0
    
    'Save the internet files at the specified folder of your hard disk.
    On Error Resume Next
    For i = 8 To LastRow
    
        'Use the given file name.
        If Not sh.Cells(i, 4) = vbNullString Then
            
            'Get the given file name.
            FilePath = sh.Cells(i, 4)
            
            'Check if the file path contains a special/illegal character.
            For j = LBound(SpecialChar) To UBound(SpecialChar)
                SpecialCharFound = InStr(1, FilePath, SpecialChar(j), vbTextCompare)
                'If an illegal character is found substitute it with a "-" character.
                If SpecialCharFound > 0 Then
                    FilePath = WorksheetFunction.Substitute(FilePath, SpecialChar(j), "-")
                End If
            Next j
            
            'Create the final file path.
            FilePath = DownloadFolder & FilePath
            
            'Check if the file path exceeds the maximum allowable characters.
            If Len(FilePath) > 255 Then
                sh.Cells(i, 5) = "ERROR"
                CountErrors = CountErrors + 1
            End If
                
        Else
            'Empty file name.
            sh.Cells(i, 5) = "ERROR"
            CountErrors = CountErrors + 1
        End If
        
        'If the file path is valid, save the file into the selected folder.
        If UCase(sh.Cells(i, 5)) <> "ERROR" Then
        
            'Try to download and save the file.
            Result = URLDownloadToFile(0, sh.Cells(i, 3), FilePath, 0, 0)
            
            'Check if the file downloaded successfully and exists.
            If Result = 0 And Not Dir(FilePath, vbDirectory) = vbNullString Then
                'Success!
                sh.Cells(i, 5) = "OK"
            Else
                'Error!
                sh.Cells(i, 5) = "ERROR"
                CountErrors = CountErrors + 1
            End If
            
        End If
        
    Next i
    On Error GoTo 0
    
    'Enable the screen.
    Application.ScreenUpdating = True
    
    'Inform the user that macro finished successfully or with errors.
    If CountErrors = 0 Then
        'Success!
        If LastRow - 7 = 1 Then
            MsgBox "The file was successfully downloaded!", vbInformation, "Done"
        Else
            MsgBox LastRow - 7 & " files were successfully downloaded!", vbInformation, "Done"
        End If
    Else
        'Error!
        If CountErrors = 1 Then
            MsgBox "There was an error with one of the files!", vbCritical, "Error"
        Else
            MsgBox "There was an error with " & CountErrors & " files!", vbCritical, "Error"
        End If
    End If
    
End Sub

