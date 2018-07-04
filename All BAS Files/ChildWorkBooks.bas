Attribute VB_Name = "ChildWorkBooks"
Sub Child_Workbooks()
Dim s As Worksheet, SH As Worksheet
Dim Master_WB As Workbook: Set Master_WB = ThisWorkbook

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
'Make New Folder
Dim MyPathFolder As String: MyPathFolder = ActiveWorkbook.Path & "\Advisors\"
Dim MyDate As String: MyDate = Now
MyDate = Format(Now, "dd-mm-yy")
MyPathFolder = InputBox("Input file path to upload sheets to...", "Upload Advisor Data", ThisWorkbook.Path & "\")
If MyPathFolder = "" Then Exit Sub

Dim Adv_BookName As String

        
'Loop Advisor Names and create worksheets in new folder
'save as csv to open with google and excel
Master_WB.Activate
Sheets("Sheet1").Activate
Dim Adv_RNG As Range: Set Adv_RNG = Sheets("Sheet1").Range("G2:G" & Rows.Count)
Dim cell As Range
For Each cell In Adv_RNG
    If cell.Value <> "" Then
        Adv_BookName = cell.Value
        Sheets("Sheet1").Cells.Copy
        Sheets.Add
        Cells(1, 1).PasteSpecial xlPasteAll
        ActiveSheet.Move
        ActiveWorkbook.SaveAs Filename:=MyPathFolder & Adv_BookName & ".csv", FileFormat:=xlCSV '
        ActiveWorkbook.Close True
        Master_WB.Activate
        Sheets("Sheet1").Activate
    End If
Next cell
Application.ScreenUpdating = True
End Sub
Sub ImportDataFromCloudFolder()
Dim folder As String
Dim source As String
Dim dest As String
Dim msg1 As String
Dim msg2 As String
Dim p As Integer
Dim s As Integer
Dim i As Long

'Prompt to set the folder path for child sheets to import
Dim FolderPath As String
On Error GoTo ErrorHandler
folder = FolderPath
msg1 = "The selected file is already in this folder."
msg2 = "was copied to"
p = 1
i = 1
    'inform user to select correct path
    FolderPath = MsgBox("Please paste the path to Advisor Folder to Import" _
    & vbNewLine & "These files will be added back to Master Folder and replaced with new data", vbOKOnly, "Import data from Advisors")
    Application.DisplayAlerts = False
    ' get the name of the file from the user
    source = Application.GetOpenFilename
    ' don’t do anything if cancelled
    If source = "False" Then Exit Sub
    ' get the total number of backslash characters "\" in the
    ' source variable’s contents
    Application.DisplayAlerts = True
    Do Until p = 0
        p = InStr(i, source, "\", 1)
        If p = 0 Then Exit Do
        s = p
        i = p + 1
    Loop
    ' create the destination filename
    dest = folder & Mid(source, s, Len(source))
    ' create a new folder with this name
    MkDir folder
    ' check if the specified file already exists in the
    ' destination folder
    If Dir(dest) <> "" Then
        MsgBox msg1
    Else
    ' copy the selected file to the C:\Abort folder
        FileCopy source, dest
        MsgBox source & " " & msg2 & " " & dest
    End If
Exit Sub
ErrorHandler:
    If Err = "75" Then
        Resume Next
    End If

    If Err = "70" Then
        MsgBox "You can’t copy an open file."
    Exit Sub
    End If
End Sub
