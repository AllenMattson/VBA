Attribute VB_Name = "ExportChild_3"
Sub UpdatedLocateChildSTART()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''UPDATE THE MASTER SHEET FROM CHILD WORKBOOKS BEFORE EXPORTING'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
ImportChild_3.UpdatedImportChildData
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''EXPORT PROCESS BELOW''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim MyPathFolder As String
Sheets("Sheet1").Activate

Application.ScreenUpdating = True
Application.DisplayAlerts = False
'delete old sheets if there is any
Dim OldSH As Worksheet
For Each OldSH In ThisWorkbook.Sheets
    If OldSH.Name <> "Sheet1" Then
        If OldSH.Name <> "Sheet2" Then
            If OldSH.Name <> "Sheet3" Then
                Application.DisplayAlerts = False
                OldSH.Delete
                Application.DisplayAlerts = True
            End If
        End If
    End If
Next OldSH
Dim Master_WB As Workbook: Set Master_WB = ThisWorkbook
Dim WWBB As Workbook
For Each WWBB In Workbooks
    If WWBB.Name <> Master_WB.Name Then WWBB.Close False
Next
Master_WB.Activate
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Find path to Child Folders
Sheets("sheet3").Activate
If Len(Range("B1")) > 2 Then
    MyPathFolder = Sheets("Sheet3").Cells(1, 2).Value
Else
    MyPathFolder = InputBox("Please Insert Folder Path to Child Sheets", "BA Folder Location", ThisWorkbook.path)
End If
If MyPathFolder = "" Then
    MsgBox "Insert Path to Import BA sheets"
    Exit Sub
End If
If Right(MyPathFolder, 1) <> "\" Then MyPathFolder = MyPathFolder & "\"
Sheets("Sheet3").Cells(1, 2).Value = MyPathFolder
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Loop to Move Data to correct workbook
Sheets("Sheet1").Activate
'Column G is the BA Sheet Column (references Child sheets/Creates Child Sheets based off column "G")
LR = Cells(Rows.Count, "G").End(xlUp).Row

For i = 2 To LR
    Cells(i, 7).Copy Destination:=Sheets("Sheet3").Range("B2")
    Sheets("Sheet1").Activate
    FilterIT (Cells(i, 7).Value)
    Master_WB.Activate
    Sheets("Sheet1").Activate
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
MoveListsOver
End Sub
Sub FilterIT(STR As String)
Dim ws As Worksheet
Dim LR As Integer
Dim wb As Workbook: Set wb = ThisWorkbook
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CREATE AUTOFILTER FOR EACH BA REFERENCE IN COLUMN "G"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Auto Filter off each Value in Column "G"
Sheets("Sheet1").Activate
Range("G1:G900").AutoFilter 1, STR
LR = Range("G" & Rows.Count).End(xlUp).Row
Application.DisplayAlerts = False
Application.ScreenUpdating = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CHECK LOOP IF BA SHEET EXISTS, SELECT THE BA SHEET
'ADD SHEET, NAME IT AFTER BA, DELETE OLD BA SHEET IF IT EXISTS>>>USED FOR CHECKING PROC
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = STR Then
            ws.Activate
        End If
    Next ws
    Sheets.Add
Start:
    On Error GoTo ErrHandle
    ActiveSheet.Name = STR
    Application.DisplayAlerts = True
        If LR < 1 Then 'No Match........TURN OFF FILTER
            Range("G1:G" & LR).AutoFilter
            Exit Sub
        Else
        'COPY FILTERED RANGE AND PASTE DATA IN BA'S SHEET
            Sheets("Sheet1").Activate
            Range("A2:BB" & LR).Copy Sheets(STR).Range("A" & Rows.Count).End(xlUp)(2)
            Range("G1:G" & LR).AutoFilter
        End If
        


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'procedure to move into correct folder
'MOVE'S BA SHEET WITH FILTERED DATA TO CLOUD FOLDER
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
MoveChildToCloudFolder (STR)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
wb.Activate: Sheets(STR).Delete: Sheets("Sheet1").Activate
Exit Sub

ErrHandle:
On Error GoTo Start
If Err.Number = 1004 Then
    ActiveSheet.Delete: Sheets(STR).Activate
    GoTo Start
Else
    If Err.Number > 0 Or Err.Number < 0 Then
        MsgBox Err.Number & vbNewLine & vbNewLine & Err.Description: Exit Sub
    Else
        'Exit Sub
    End If
End If
End Sub
Sub MoveChildToCloudFolder(STR As String)
Application.DisplayAlerts = False
Application.EnableEvents = True
'MOVE HEADERS FROM SHEET1 TO THE BA SHEET
Sheets("Sheet1").Activate: Rows("1:1").Copy Sheets(STR).Rows("1:1")

Dim FullPath As String
Sheets("Sheet3").Activate: FullPath = Cells(1, 2).Value
Cells(2, 2).Value = STR
Sheets(STR).Activate
Columns.AutoFit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''*************IMPORTANT LINE OF THIS PROCEDURE**************'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Export BA Sheet to New Workbook''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets(STR).Move
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim wb As Workbook
    Set wb = ActiveWorkbook
    'fileSaveName = Application.GetSaveAsFilename(InitialFileName:=STR, _
    'filefilter:="Excel files , *.xlsx")
With wb
''''''''''''''''''''''''''''''''''''''''''''
'do not show secret data on child sheets
'with child activated delete columns
''''''''''''''''''''''''''''''''''''''''''''
Columns("BC").EntireColumn.Delete
'Columns("AX:AZ").EntireColumn.Delete
'Columns("C:S").EntireColumn.Delete
'Columns("A").EntireColumn.Delete
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PROMPT TO SAVE WORKBOOK
'CURRENTLY CLOSE WITHOUT SAVING
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Application.DisplayAlerts = False
        'If fileSaveName <> "False" Then
    .SaveAs FullPath & STR & ".xlsx" 'fileSaveName
    .Close False
        'Else
    '.Close False
    Exit Sub
       ' End If
End With
    
ActiveWorkbook.SaveAs FullPath & STR & ".xlsx", FileFormat:=vbNormal
ActiveWorkbook.Close True ', FullPath & STR 'SaveAs FullPath & STR.xls
Application.DisplayAlerts = True
End Sub
Sub MoveListsOver()
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.Calculation = xlCalculationManual


Dim wb As Workbook: Set wb = ActiveWorkbook
Dim Lsh As Worksheet: Set Lsh = Sheets("Sheet2")
Sheets("Sheet3").Activate

'Locate BA folder path
Dim BAfolderPath As String: BAfolderPath = Range("B1").Value

Dim FileName As String
Dim PathAndName As String
Dim n As Name
FileName = Dir(BAfolderPath & "*.xlsx")
Do While Len(FileName) <> 0
    Dim BAWB As Workbook
    Dim Osh As Worksheet
        If Left(FileName, 4) <> ".xlsx" Then 'current dir
            PathAndName = BAfolderPath & FileName
            'Open BA WorkBook
            Set BAWB = Workbooks.Open(PathAndName)
            'Delete Old Drop Downs if one Exists
            For Each Osh In BAWB.Worksheets
                If Osh.Visible = xlSheetHidden Then
                    If Osh.Name = "Sheet2" Then
                        Osh.Visible = True
                        Osh.Delete
                    End If
                Else
                    For Each n In Osh.Names
                        n.Delete
                    Next
                End If
            Next Osh
Calculate
DoEvents
            Application.ScreenUpdating = True
            
            'Transfer Sheet2 to BA Sheet
            wb.Activate: Lsh.Select
            Lsh.Copy After:=Workbooks(FileName).Sheets(1)
            'Call Macro to Remove Names and Ranges
            'Insert New Drop Downs and Hide DropDown Sheet
            
            'Close BA WorkBook and Hide Drop list sheet
            BAWB.Activate
            For Each n In ActiveSheet.Names
                n.Delete
            Next
            SetRangeNames_3.PermissionGranted
            Sheets("Sheet2").Visible = xlHidden
            BAWB.Close True
            wb.Activate
        End If
    FileName = Dir()
Loop
'Reset Macro Optimization Settings
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub
Function IsMac() As Boolean
#If Mac Then
    IsMac = True
#ElseIf Win32 Or Win64 Then
    IsMac = False
#End If
End Function

Function FileExists(ByVal AFileName As String) As Boolean
    On Error GoTo Catch

    FileSystem.FileLen AFileName

    FileExists = True

    GoTo Finally

Catch:
        FileExists = False
Finally:
End Function


