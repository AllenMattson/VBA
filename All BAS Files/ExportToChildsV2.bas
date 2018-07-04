Attribute VB_Name = "ExportToChildsV2"
Sub LocateChild()


Application.ScreenUpdating = False
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

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Find path to Child Folders
Dim MyPathFolder As String: MyPathFolder = InputBox("Input file path to upload sheets to...", "Upload Advisor Data", ThisWorkbook.path)
If MyPathFolder = "" Then
MsgBox "Enter the folder path to the Child Sheets"
Exit Sub
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''check if pc or mac'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'make sure there is a backslash or colon to end file path
    If Right(MyPathFolder, 1) <> "\" Then
        If Right(MyPathFolder, 1) <> ":" Then
            If IsMac = True Then
                MyPathFolder = MyPathFolder & ":"
                Master_WB.Sheets("Sheet3").Cells(3, 2).Clear
                Master_WB.Sheets("Sheet3").Cells(3, 2).Value = "Mac"
            Else
                MyPathFolder = MyPathFolder & "\"
                Sheets("Sheet3").Activate
                Sheets("Sheet3").Cells(3, 2).Clear
                Sheets("Sheet3").Cells(3, 2).Value = "PC"
            End If
        End If
    End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Create Loop to Move Data to correct workbook
Sheets("Sheet1").Activate
lr = Sheets("Sheet1").Cells(Rows.Count, 7).End(xlUp).Row

For i = 2 To lr
    FilterIT (Cells(i, 7).Value)
    Sheets("Sheet1").Activate
Next i
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
        If lr < 1 Then 'No Match
            Range("G1:G" & lr).AutoFilter
            Exit Sub
        Else
            Sheets("Sheet1").Activate
            Range("A2:BB" & lr).Copy Sheets(STR).Range("A" & Rows.Count).End(xlUp)(2)
            Range("G1:G" & lr).AutoFilter
        End If
        

Sheets("Sheet1").Activate
Range("G1:G" & lr).AutoFilter


'procedure to move into correct folder
MoveChildToCloudFolder (STR)
wb.Activate: Sheets("Sheet1").Activate
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
        Exit Sub
    End If
End If
End Sub
Sub MoveChildToCloudFolder(STR As String)
Application.DisplayAlerts = False
Sheets("Sheet1").Activate: Rows("1:1").Copy Sheets(STR).Rows("1:1")
Dim FullPath As String
Sheets("Sheet3").Activate
FullPath = Cells(1, 2).Value
Sheets(STR).Activate
Sheets(STR).Move
ActiveWorkbook.Close True, FullPath & STR 'SaveAs FullPath & STR.xls
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


