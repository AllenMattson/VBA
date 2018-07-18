VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Export Charts"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   HelpContextID   =   10510
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim TheDir As String
    ExportFormatCombo.List = Array("GIF", "JPEG", "TIF", "PNG")
    TheDir = ActiveWorkbook.path
    If TheDir = "" Then TheDir = Application.DefaultFilePath
    TextBoxOutputDirectory.Text = TheDir
    
    If GetSetting(APPNAME, "Settings", "RememberSettings", 1) = 1 Then
        ExportFormatCombo.ListIndex = GetSetting(APPNAME, "Settings", "ExportFormatCombo", 0)
        TextBoxOutputDirectory.Text = GetSetting(APPNAME, "Settings", "TextBoxOutputDirectory", TheDir)
        WarnCheckBox.Value = GetSetting(APPNAME, "Settings", "WarnCheckBox", True)
    End If
End Sub

Private Sub ScrollToChartButton_Click()
    Dim i As Long
    Dim ChartCell As Range
    On Error Resume Next
    If TypeName(ActiveSheet) = "Chart" Then
        For i = 0 To ChartList.ListCount - 1
            If ChartList.Selected(i) Then
                Sheets(ChartData(0, i)).Activate
                Exit For
            End If
        Next i
    Else
        For i = 0 To ChartList.ListCount - 1
            If ChartList.Selected(i) Then
                Set ChartCell = ActiveSheet.ChartObjects(ChartData(0, i)).TopLeftCell
                Application.GoTo ChartCell, True
                Exit For
            End If
        Next i
    End If
    On Error GoTo 0
End Sub


Private Sub ChartList_Change()
    Dim SelCnt As Long
    Dim i As Long
    SelCnt = 0
    For i = 0 To ChartList.ListCount - 1
        If ChartList.Selected(i) Then SelCnt = SelCnt + 1
    Next i
    If SelCnt > 1 Then
        RenameButton.Enabled = False
        ScrollToChartButton.Enabled = False
    Else
        RenameButton.Enabled = True
        ScrollToChartButton.Enabled = True
    End If
End Sub

Private Sub RenameButton_Click()
    Dim i As Long
    Me.Hide
    For i = 0 To ChartList.ListCount - 1
        If ChartList.Selected(i) Then
            SelectedChartIndex = i
            Exit For
        End If
    Next i
    With UserForm2
      .StartUpPosition = 0
      .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
      .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
      .Show
    End With
End Sub


Sub OKButton_Click()
    Dim i As Integer, Ans As Integer
    Dim ExportFormat As String, fName As String
    Dim CurrentChart As Chart
    Dim OutputDirectory As String
    Dim ErrorCount As Integer
    Dim SaveTheChart As Boolean
    Dim Msg As String
    Dim SelCnt As Long
    
    If TextBoxOutputDirectory.Text = "" Then Call ChangeDirectoryButton_Click
    If TextBoxOutputDirectory.Text = "" Then Exit Sub
    OutputDirectory = TextBoxOutputDirectory.Text

'   Make sure output directory is writable
    If Not CanWriteToDirectory(OutputDirectory) Then
        MsgBox "Error..." & vbCrLf & vbCrLf & "Cannot write to " & OutputDirectory & vbCrLf & vbCrLf & "The directory may not exist, or it may be write-protected. ", vbCritical, APPNAME
        Exit Sub
    End If

'   Make sure at least one is selected
    SelCnt = 0
    For i = 0 To ChartList.ListCount - 1
        If ChartList.Selected(i) Then SelCnt = SelCnt + 1
    Next i
    If SelCnt = 0 Then
        MsgBox "No charts are selected.", vbInformation, APPNAME
        Exit Sub
    End If

    Me.Hide
    DoEvents
    Application.ScreenUpdating = False
    ErrorCount = 0
    For i = 0 To ChartList.ListCount - 1
        If ChartList.Selected(i) Then
            SaveTheChart = True
            Application.StatusBar = "Processing " & ChartList.List(i)
            If TypeName(ActiveSheet) = "Worksheet" Then
                Set CurrentChart = ActiveSheet.ChartObjects(ChartList.List(i)).Chart
            Else 'Chartsheet
                Set CurrentChart = ActiveWorkbook.Sheets(ChartList.List(i))
            End If
            ExportFormat = ExportFormatCombo.Value
            fName = OutputDirectory & Application.PathSeparator & ChartData(1, i)
            If UserForm1.WarnCheckBox Then
                If Dir(fName) <> "" Then
                    Application.ScreenUpdating = True
                    Ans = MsgBox(fName & vbCrLf & vbCrLf & "This file already exists. Do you want to replace the existing copy?", vbQuestion + vbYesNo, APPNAME)
                    If Ans = vbNo Then SaveTheChart = False
                    Application.ScreenUpdating = False
                End If
            End If
            On Error Resume Next
            If SaveTheChart Then CurrentChart.Export FileName:=fName, FilterName:=ExportFormat
            If Err <> 0 Then
                ErrorCount = ErrorCount + 1
                Kill fName
            End If
            On Error GoTo 0
        End If
    Next i
    Application.StatusBar = False
    Application.ScreenUpdating = True
    SaveSetting APPNAME, "Settings", "ExportFormatCombo", ExportFormatCombo.ListIndex
    SaveSetting APPNAME, "Settings", "TextBoxOutputDirectory", TextBoxOutputDirectory.Text
    SaveSetting APPNAME, "Settings", "WarnCheckBox", WarnCheckBox.Value

    If ErrorCount <> 0 Then
        Msg = ""
        Msg = vbCrLf & vbCrLf & "It's possible that the " & ExportFormat & " graphics converter is not installed on your system."
        MsgBox "An error occured." & vbCrLf & vbCrLf & ErrorCount & " charts could not be exported." & Msg, vbCritical, APPNAME
    End If
    GoToOriginalCell
    Unload Me
End Sub

Private Function CanWriteToDirectory(d) As Boolean
'   Returns True if the directory can be written to
    On Error Resume Next
    Open d & Application.PathSeparator & "testfile" For Output As #1
    If Err.Number = 0 Then
        CanWriteToDirectory = True
        Close #1
        Kill d & Application.PathSeparator & "testfile"
    Else
        CanWriteToDirectory = False
    End If
    On Error Resume Next
End Function

Private Sub CancelButton_Click()
    GoToOriginalCell
    Unload Me
End Sub

Private Sub GoToOriginalCell()
    On Error Resume Next
    If TypeName(ActiveSheet) = "Worksheet" Then
        ActiveWindow.ScrollRow = UserRow
        ActiveWindow.ScrollColumn = UserCol
    End If
    On Error GoTo 0
End Sub
Private Sub ChangeDirectoryButton_Click()
'   Gets a new directory
    Dim NewDir As String
    Dim fd As FileDialog
    Dim ScrollToSetting As Boolean, RenameSetting As Boolean
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        If .Show = -1 Then
           NewDir = .SelectedItems(1)
        Else
            NewDir = ""
        End If
    End With
    
    If NewDir <> "" Then
        If Right(NewDir, 1) = "\" Then NewDir = Left(NewDir, Len(NewDir) - 1)
        With TextBoxOutputDirectory
            .Text = NewDir
            .ControlTipText = .Text
        End With
    End If
    
    
End Sub

Private Sub ExportFormatCombo_Change()
    Dim i As Integer, Extension As String
    Dim SaveSelected()
    ReDim SaveSelected(ChartList.ListCount)
    
    If Me.Visible Then
        For i = 0 To ChartList.ListCount - 1
            SaveSelected(i) = ChartList.Selected(i)
            Extension = LCase(ExportFormatCombo.Value)
            If Extension = "jpeg" Then Extension = "jpg"
            ChartData(1, i) = LCase(Replace(ChartData(0, i), " ", "_") & "." & Extension)
        Next i
        ChartList.Column = ChartData
        For i = 0 To ChartList.ListCount - 1
            ChartList.Selected(i) = SaveSelected(i)
        Next
    End If
End Sub

Private Sub HelpButton_Click()
    Call ShowHelp
End Sub




