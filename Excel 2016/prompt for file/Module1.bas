Attribute VB_Name = "Module1"
Option Explicit

Sub GetImportFileName()
    Dim Filt As String
    Dim FilterIndex As Long
    Dim Title As String
    Dim FileName As Variant
    
'   Set up list of file filters
    Filt = "Text Files (*.txt),*.txt," & _
           "Lotus Files (*.prn),*.prn," & _
           "Comma Separated Files (*.csv),*.csv," & _
           "ASCII Files (*.asc),*.asc," & _
           "All Files (*.*),*.*"

'   Display *.* by default
    FilterIndex = 5

'   Set the dialog box caption
    Title = "Select a File to Import"

'   Get the file name
    FileName = Application.GetOpenFilename _
        (FileFilter:=Filt, _
         FilterIndex:=FilterIndex, _
         Title:=Title)

'   Exit if dialog box canceled
    If FileName <> False Then
    '   Display full path and name of the file
        MsgBox "You selected " & FileName
    Else
        MsgBox "No file was selected."
    End If
   
End Sub


Sub GetImportFileName2()
    Dim Filt As String
    Dim FilterIndex As Long
    Dim FileName As Variant
    Dim Title As String
    Dim i As Long
    Dim Msg As String
'   Set up list of file filters
    Filt = "Text Files (*.txt),*.txt," & _
            "Lotus Files (*.prn),*.prn," & _
            "Comma Separated Files (*.csv),*.csv," & _
            "ASCII Files (*.asc),*.asc," & _
            "All Files (*.*),*.*"
'   Display *.* by default
    FilterIndex = 5

'   Set the dialog box caption
    Title = "Select a File to Import"

'   Get the file name
    FileName = Application.GetOpenFilename _
        (FileFilter:=Filt, _
         FilterIndex:=FilterIndex, _
         Title:=Title, _
         MultiSelect:=True)

    If IsArray(FileName) Then
    '   Display full path and name of the files
        For i = LBound(FileName) To UBound(FileName)
            Msg = Msg & FileName(i) & vbNewLine
        Next i
        MsgBox "You selected:" & vbNewLine & Msg
    Else
    '   Exit if dialog box canceled
        MsgBox "No file was selected."
    End If
    
End Sub


Sub GetAFolder()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"

        .Title = "Select a location for the backup"
        .Show
        If .SelectedItems.Count = 0 Then
            MsgBox "Canceled"
        Else
            MsgBox .SelectedItems(1)
        End If
    End With
End Sub


