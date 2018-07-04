Attribute VB_Name = "ImportCSVfiles"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function OpenProcess Lib "kernel32" _
        (ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    
    Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As Long, _
        lpExitCode As Long) As Long
#Else
    Private Declare Function OpenProcess Lib "kernel32" _
        (ByVal dwDesiredAccess As Long, _
        ByVal bInheritHandle As Long, _
        ByVal dwProcessId As Long) As Long
    
    Private Declare Function GetExitCodeProcess Lib "kernel32" _
        (ByVal hProcess As Long, _
        lpExitCode As Long) As Long
#End If


Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STILL_ACTIVE = &H103
Public ChildSheetLocation As String


Public Sub ShellAndWait(ByVal PathName As String, Optional WindowState)
    Dim hProg As Long
    Dim hProcess As Long, ExitCode As Long
    'fill in the missing parameter and execute the program
    If IsMissing(WindowState) Then WindowState = 1
    hProg = Shell(PathName, WindowState)
    'hProg is a "process ID under Win32. To get the process handle:
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, hProg)
    Do
        'populate Exitcode variable
        GetExitCodeProcess hProcess, ExitCode
        DoEvents
    Loop While ExitCode = STILL_ACTIVE
End Sub


Sub SaveAsXLS()



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This Module starts with SaveAsXLS and converts the advisor data back
'to a csv file to compile together and import into master
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''





'Take new advisor data and convert to csv files and compile into xlsx doc
'Update Master Sheet with new data
'Remove old advisor files
Dim wb As Workbook
Dim sh As Worksheet
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog
Dim nameWB
'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = False
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then Exit Sub

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls"

'Target Path with Ending Extention
myFile = Dir(myPath & myExtension)

On Error GoTo ErrHandle
'Loop through each Excel file in folder
  Do While myFile <> ""
    If Right(myFile, 1) = "m" Then GoTo ErrHandle
    If Right(myFile, 1) = "x" Then GoTo ErrHandle
    
    'Set variable equal to opened workbook
    Set wb = Workbooks.Open(FileName:=myPath & myFile)

        Application.DisplayAlerts = False
        nameWB = myPath & Left(myFile, InStr(1, myFile, ".") - 1) & ".csv"
        ActiveWorkbook.SaveAs FileName:=nameWB, FileFormat:=xlCSV
        ActiveWorkbook.Close savechanges:=False
        'Get next file name
        myFile = Dir
        Application.DisplayAlerts = True
  Loop
  
  
  
ErrHandle:
    If Err.Number < 0 Then
        Merge_CSV_Files
    Else
        If Err.Number > 0 Then
            MsgBox "Error Number: " & Err.Number & vbNewLine & Err.Description & vbNewLine & "Source: " & Err.source
            Exit Sub
        Else
            'nothing
        End If
    End If
    

'Reset Macro Optimization Settings
  Application.ScreenUpdating = True
  Application.EnableEvents = True
  Application.Calculation = xlCalculationAutomatic

ImportCSVfiles.Merge_CSV_Files



End Sub
Sub Merge_CSV_Files()
Dim MasterWB As Workbook: Set MasterWB = ThisWorkbook
Dim Mpath As String: Mpath = MasterWB.Path & "\"
    Dim UpdatedPath As String
    Dim BatFileName As String
    Dim TXTFileName As String
    Dim XLSFileName As String
    Dim FileExtStr As String
    Dim FileFormatNum As Long
    Dim DefPath As String
    Dim wb As Workbook
    Dim oApp As Object
    Dim oFolder
    Dim foldername

    'Create two temporary file names
    BatFileName = Environ("Temp") & _
            "\CollectCSVData" & Format(Now, "dd-mm-yy-h-mm-ss") & ".bat"
    TXTFileName = Environ("Temp") & _
            "\AllCSV" & Format(Now, "dd-mm-yy-h-mm-ss") & ".txt"

    'Folder where you want to save the Excel file
    DefPath = InputBox("Where would you like to save the updated Master Sheet?", "Advisor Import Data", ThisWorkbook.Path) 'Application.DefaultFilePath
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If

    'Set the extension and file format
    If Val(Application.Version) < 12 Then
        'You use Excel 97-2003
        FileExtStr = ".xls": FileFormatNum = -4143
    Else
        'You use Excel 2007 or higher
        FileExtStr = ".xlsx": FileFormatNum = 51
        'If you want to save as xls(97-2003 format) in 2007 use
        'FileExtStr = ".xls": FileFormatNum = 56
    End If

    'Name of the Excel file with a date/time stamp
    XLSFileName = DefPath & "Master " & _
                  Format(Now, "dd-mmm-yyyy h-mm-ss") & FileExtStr

    'Browse to the folder with CSV files
    Set oApp = CreateObject("Shell.Application")
    Set oFolder = oApp.BrowseForFolder(0, "Select folder with advisor files to compile", 512)
    If Not oFolder Is Nothing Then
        foldername = oFolder.Self.Path
        If Right(foldername, 1) <> "\" Then
            foldername = foldername & "\"
        End If

        'Create the bat file
        Open BatFileName For Output As #1
        Print #1, "Copy " & Chr(34) & foldername & "*.csv" _
                & Chr(34) & " " & TXTFileName
        Close #1

        'Run the Bat file to collect all data from the CSV files into a TXT file
        ShellAndWait BatFileName, 0
        If Dir(TXTFileName) = "" Then
            MsgBox "There are no appropriate files in this folder to import" & vbNewLine & vbNewLine & "Please recapture and select correct BA folder"
            Kill BatFileName
            Exit Sub
        End If

        'Open the TXT file in Excel
        Application.ScreenUpdating = False
        Workbooks.OpenText FileName:=TXTFileName, Origin:=xlWindows, StartRow _
                :=2, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, Comma:=True, _
                Space:=False, Other:=False

        'Save text file as a Excel file
        Set wb = ActiveWorkbook
        wb.ActiveSheet.Name = "Sheet1"
        Application.DisplayAlerts = False
        UpdatedPath = XLSFileName
        wb.SaveAs FileName:=XLSFileName, FileFormat:=FileFormatNum
        Application.DisplayAlerts = True

        wb.Close savechanges:=False
        MsgBox "You find the Excel file here: " & vbNewLine & XLSFileName

        'Delete the bat and text file you temporary used
        Kill BatFileName
        Kill TXTFileName

        Application.ScreenUpdating = True
    End If
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    'send compiled updated workbook to master sheet
'    Application.Open (UpdatedPath)
'    Worksheets(1).Cells(1, 1).CurrentRegion.Copy
'    MasterWB.Activate
'    Sheets("Sheet1").Cells(1, 1).PasteSpecial xlPasteAll
    
    
    '''''''''''''''''''''''''''''''''''''''''''''''''
    'DELETE OLD DATA NOW THAT IT IS IMPORTED
    'Delete all files and subfolders
'Be sure that no file is open in the folder
    Dim FSO As Object
    Dim myPath As String

    Set FSO = CreateObject("scripting.filesystemobject")

    myPath = foldername

    If Right(myPath, 1) = "\" Then
        myPath = Left(myPath, Len(myPath) - 1)
    End If

    If FSO.FolderExists(myPath) = False Then
        'MsgBox MyPath & " doesn't exist"
        Exit Sub
    End If

    MoveNewData (XLSFileName)
    
    On Error Resume Next
    'Delete files
    FSO.deletefile myPath & "\*.csv", True
    On Error GoTo 0
End Sub
Sub MoveNewData(UpdatedPath As String)
    Dim MasterWB As Workbook: Set MasterWB = ThisWorkbook
    Dim temp As Workbook: Set temp = Workbooks.Open(UpdatedPath)
    Dim B As Integer, LRR As Integer: LRR = Cells(Rows.Count, 6).End(xlUp).Row
    For B = LRR To 2 Step -1
        If Cells(B, 1).Value = "Date" Then Rows(B).Delete shift:=xlUp: DoEvents
    Next B
    
    'copy new data
    Cells(1, 1).CurrentRegion.Copy
    
    'move to master
    MasterWB.Activate: Sheets("Sheet1").Activate: Cells(2, 1).PasteSpecial xlPasteValues: Cells(1, 1).Select
    temp.Close False
End Sub


