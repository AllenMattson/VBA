Attribute VB_Name = "Module1"
Sub GetImportFileName()
    Dim Filt As String
    Dim FilterIndex As Integer
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
    If FileName = False Then
        MsgBox "No file was selected."
        Exit Sub
    End If
   
'   Display full path and name of the file
    MsgBox "You selected " & FileName
End Sub


Sub GetImportFileName2()
    Dim Filt As String
    Dim FilterIndex As Integer
    Dim FileName As Variant
    Dim Title As String
    Dim i As Integer
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

'   Exit if dialog box canceled
    If Not IsArray(FileName) Then
        MsgBox "No file was selected."
        Exit Sub
    End If
   
'   Display full path and name of the files
    For i = LBound(FileName) To UBound(FileName)
        Msg = Msg & FileName(i) & vbCrLf
    Next i
    MsgBox "You selected:" & vbCrLf & Msg
End Sub


