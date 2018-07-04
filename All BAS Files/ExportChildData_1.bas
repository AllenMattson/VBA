Attribute VB_Name = "ExportChildData_1"
Sub LocateChildSTART()


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
Sheets("Sheet3").Activate
Dim MyPathFolder As String: MyPathFolder = InputBox("Input file path to upload sheets to...", "Upload Advisor Data", ThisWorkbook.path & "\")
Sheets("Sheet3").Activate
Debug.Print MyPathFolder
DoEvents
Sheets("Sheet3").Cells(1, 2).Value = MyPathFolder
'If Cells(1, 2).Value = "" Then Exit Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''check if pc or mac'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'make sure there is a backslash or colon to end file path
    If Right(MyPathFolder, 1) <> "\" Then
        If Right(MyPathFolder, 1) <> ":" Then

            If IsMac = True Then
                Cells(1, 2).Clear
                Cells(1, 2).Value = MyPathFolder & ":"
                Cells(3, 2).Clear
                Cells(3, 2).Value = "Mac"
            
            Else
                Cells(1, 2).Clear
                Cells(1, 2).Value = MyPathFolder & "\"
                Cells(3, 2).Clear
                Cells(3, 2).Value = "PC"
            End If
        End If
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Create Loop to Move Data to correct workbook
Sheets("Sheet1").Activate
lr = Cells(Rows.Count, 7).End(xlUp).Row

For i = 2 To lr
    Cells(i, 7).Copy Destination:=Sheets("Sheet3").Range("B2")
    Sheets("Sheet1").Activate
    FilterIT (Cells(i, 7).Value)
    Master_WB.Activate
    Sheets("Sheet1").Activate
Next i
Sheets("Sheet3").Activate: Range("B1:B3").ClearFormats: Range("B2").Validation.Delete

End Sub
Sub FilterIT(STR As String)
    'Dim Str As String
    Dim ws As Worksheet
    Dim lr As Integer
    Dim wb As Workbook: Set wb = ThisWorkbook
Sheets("Sheet1").Activate
    Range("G1:G900").AutoFilter 1, STR
    lr = Range("G" & Rows.Count).End(xlUp).Row
    Application.DisplayAlerts = False
    Application.ScreenUpdating = True
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = STR Then
            ws.Activate
        End If
    Next ws
    Sheets.add
Start:
    On Error GoTo ErrHandle
    ActiveSheet.Name = STR
    Application.DisplayAlerts = True
        If lr < 1 Then 'No Match
            Range("G1:G" & lr).AutoFilter
            Exit Sub
        Else
            Sheets("Sheet1").Activate
            Range("A2:BB" & lr).Copy Sheets(STR).Range("A" & Rows.Count).End(xlUp)(2)
            Range("G1:G" & lr).AutoFilter
        End If
        



'procedure to move into correct folder
MoveChildToCloudFolder (STR)
'NewExportProc (STR) 'MoveChildToCloudFolder (STR)
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
Sheets("Sheet1").Activate: Rows("1:1").Copy Sheets(STR).Rows("1:1")




Dim FullPath As String
Sheets("Sheet3").Activate: FullPath = Cells(1, 2).Value
Cells(2, 2).Value = STR
Sheets(STR).Activate
Columns.AutoFit
Sheets(STR).Move
Dim wb As Workbook
    Set wb = ActiveWorkbook
     
    'fileSaveName = Application.GetSaveAsFilename(InitialFileName:=STR, _
    'filefilter:="Excel files , *.xlsx")
     
    With wb
    Application.DisplayAlerts = False
        'If fileSaveName <> "False" Then
             
            .SaveAs FullPath & STR & ".xlsx" 'fileSaveName
            .Close False
        'Else
            .Close False
            Exit Sub
       ' End If
    End With
    
ActiveWorkbook.SaveAs FullPath & STR & ".xlsx", FileFormat:=vbNormal
ActiveWorkbook.Close True ', FullPath & STR 'SaveAs FullPath & STR.xls
Application.DisplayAlerts = True
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

Sub NewExportProc(STR As String)
Dim MainWB As Workbook: Set MainWB = ThisWorkbook
Sheets("Sheet1").Activate: Rows("1:1").Copy Sheets(STR).Rows("1:1")
Sheets("Sheet3").Activate
If Left(Range("B1").Value, 1) <> "\" And Range("B3").Value = "PC" Then
    'Range("B1").Value = Range("B1").Value & "\"
    
ElseIf Left(Range("B1").Value, 1) <> ":" And Range("B3").Value = "MAC" Then
    'Range("B1").Value = Range("B1").Value & ":"
Else
End If
    
Range("B2").Value = STR

    
Dim path As String: path = Range("B1").Value

MainWB.Activate

MainWB.Activate
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
Application.DisplayAlerts = False
Application.ScreenUpdating = False

    If ws.Name <> "Sheet1" And ws.Name <> "Sheet2" And ws.Name <> "Sheet3" Then
        If Left(ws.Name, 5) = "Sheet" Then
        Debug.Print ws.Name
            If ws.Name <> "Sheet3" And ws.Name <> "Sheet2" And ws.Name <> "Sheet1" Then ws.Delete
        Else
        Dim wb As Workbook
        Set wb = ws.Application.Workbooks.add
        ws.Copy Before:=wb.Sheets(1): Sheets("Sheet1").Delete
        wb.SaveAs path & ws.Name, FileFormat:=xlOpenXMLWorkbook ' Excel.XlFileFormat.xlOpenXMLWorkbook
        wb.Close True
        Kill wb
        Set wb = Nothing
        MainWB.Activate: ws.Delete
        End If
    Else
        MainWB.Activate
    End If
    MainWB.Activate
Next ws
End Sub
