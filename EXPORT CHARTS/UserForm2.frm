VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub HelpButton_Click()
    Call ShowHelp
End Sub

Private Sub UserForm_Initialize()
    If TypeName(ActiveSheet) = "Chart" Then
        Me.Caption = "Rename Chart Sheet"
        Label1.Caption = "New name for the Chart sheet:"
        Label2.Caption = "The file name is derived from the Chart sheet's name. If you change the Chart sheet's name, the file name will also change."
    Else
        Me.Caption = "Rename Chart Object"
        Label1.Caption = "New name for the Chart object:"
        Label2.Caption = "The file name is derived from the Chart object's name. If you change the Chart object's name, the file name will also change."
    End If
    With TextBoxNewName
        .Text = ChartData(0, SelectedChartIndex)
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub OKButton_Click()
    Dim i As Long
    Dim Extension As String

'   Check for duplicate name (Excel allows it, but it would mess up the file saving)
    For i = 0 To UBound(ChartData, 2)
        If UCase(TextBoxNewName.Text) = UCase(ChartData(0, i)) Then
            MsgBox TextBoxNewName.Text & vbCrLf & vbCrLf & "That name is already in use.", vbInformation, APPNAME
            Exit Sub
        End If
    Next i

    If TypeName(ActiveSheet) = "Worksheet" Then
        On Error Resume Next
        ActiveSheet.ChartObjects(SelectedChartIndex + 1).Name = TextBoxNewName.Text
        If Err.Number <> 0 Or TextBoxNewName.Text = "" Then
            MsgBox "Cannot rename the chart.", vbInformation, APPNAME
            With TextBoxNewName
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End With
            On Error GoTo 0
            Exit Sub
        Else
            On Error GoTo 0
            ChartData(0, SelectedChartIndex) = TextBoxNewName.Text
            With UserForm1
                For i = 0 To .ChartList.ListCount - 1
                    Extension = LCase(.ExportFormatCombo.Value)
                    If Extension = "jpeg" Then Extension = "jpg"
                    ChartData(1, i) = LCase(Replace(ChartData(0, i), " ", "_") & "." & Extension)
                Next i
                .ChartList.Column = ChartData
                .ChartList.Selected(SelectedChartIndex) = True
            End With
            Unload Me
            UserForm1.Show
        End If
    Else ' chart sheet
        On Error Resume Next
        ActiveWorkbook.Charts(SelectedChartIndex + 1).Name = TextBoxNewName.Text
        If Err.Number <> 0 Then
            MsgBox "Cannot rename the sheet.", vbInformation, APPNAME
            With TextBoxNewName
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End With
            On Error GoTo 0
            Exit Sub
        Else
            On Error GoTo 0
            ChartData(0, SelectedChartIndex) = TextBoxNewName.Text
            With UserForm1
                For i = 0 To .ChartList.ListCount - 1
                    Extension = LCase(.ExportFormatCombo.Value)
                    If Extension = "jpeg" Then Extension = "jpg"
                    ChartData(1, i) = LCase(Replace(ChartData(0, i), " ", "_") & "." & Extension)
                Next i
                .ChartList.Column = ChartData
                .ChartList.Selected(SelectedChartIndex) = True
            End With
            Unload Me
            UserForm1.Show
        End If
    End If
End Sub

Private Sub CancelButton_Click()
    Unload Me
    UserForm1.Show
End Sub

Private Function RemoveExtension(fn) As String
'   Removes the extension from a filename
    Dim i As Long
    Dim LastDot As Long
    For i = Len(fn) To 1 Step -1
        If Mid(fn, i, 1) = "." Then
            LastDot = i
            Exit For
        End If
    Next i
    RemoveExtension = Left(fn, LastDot - 1)
End Function

