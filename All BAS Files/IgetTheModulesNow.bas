Attribute VB_Name = "ListModules_Procedures"
Sub LoopAllExcelFilesInFolder()
'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them
'SOURCE: www.TheSpreadsheetGuru.com
Dim MainBook As Workbook: Set MainBook = ThisWorkbook
Dim wb As Workbook
Dim myPath As String
Dim MyFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog

'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'Retrieve Target Folder Path From User
  Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

    With FldrPicker
      .Title = "Select A Target Folder"
      .AllowMultiSelect = True
        If .Show <> -1 Then GoTo NextCode
        myPath = .SelectedItems(1) & "\"
    End With

'In Case of Cancel
NextCode:
  myPath = myPath
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*.xls*"

'Target Path with Ending Extention
  MyFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While MyFile <> ""
  'MainBook.Activate: PrintModuleProcedureSheet
    'Set variable equal to opened workbook
      Set wb = Workbooks.Open(FileName:=myPath & MyFile)
    
    'Ensure Workbook has opened before moving on to next line of code
      DoEvents
    'Import the Modules
    Import_VBA (MyFile)
        
    'Change First Worksheet's Background Fill Blue
      'wb.Worksheets(1).Range("A1:Z1").Interior.Color = RGB(51, 98, 174)
    
    'Save and Close Workbook
      wb.Close SaveChanges:=False
      
    'Ensure Workbook has closed before moving on to next line of code
      DoEvents

    'Get next file name
      MyFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
Sub Import_VBA(WBName As String)

    Dim VBc As Variant
    Dim exportFolder As String, VBcExt As String, testFile As String
    Dim newWB As Workbook

    testFile = WBName
    exportFolder = "C:\Users\Allen\Desktop\BuddaKing\"
    Set newWB = Workbooks(testFile)

    '''''  Test VBA protection
    On Error Resume Next
    If newWB.VBProject.Protection <> 0 Then

        If Err.Number = 1004 Then
            Err.Clear
            MsgBox "VBA Project Object Model is protected in " & newWB.Name & vbCrLf _
            & vbCrLf & "Please remove VBA protection in Trust Center to continue.", _
            vbExclamation + vbOKOnly, "Error"

            Set newWB = Nothing
            Exit Sub
        Else
            MsgBox Err.Number & ": " & Err.Description, vbExclamation, "Error"
            Set newWB = Nothing
            Err.Clear
            Exit Sub
        End If

    End If

    ''''' Add Interface App Components
    For Each VBc In CreateObject("Scripting.FileSystemObject").GetFolder(exportFolder).Files
        Select Case LCase(Right(VBc.Name, 4))
        Case ".bas", ".frm", ".cls", ".frx"
            newWB.VBProject.VBComponents.Import exportFolder & "\" & VBc.Name
        End Select
    Next VBc
End Sub
