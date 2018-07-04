Attribute VB_Name = "LoopFolderChangeExtension"
Sub LoopThroughFilesChangeCSVtoXLS()
Dim strFile As String, StrDate As String, StrDir As String
Dim wb As Workbook
StrDir = "C:\Users\Allen\Desktop\Housing VBA\Zipped Files\zip\zip\"
'strFile = ActiveWorkbook.Path & "\*" 'Dir("C:\Users\Allen\Desktop\Housing VBA\FolderTimeStamp_*") '
StrDate = Format(Now, "dd-mm-yy") '& "\" & cell & "\"
'StrDir = StrDir & StrDate
strFile = Dir(StrDir & "*.csv")
Application.ScreenUpdating = False
Do While Len(strFile) > 0
    If Right(strFile, 4) = ".csv" Then
        Set wb = Workbooks.Open(Filename:=StrDir & strFile, local:=True)
        wb.SaveAs Replace(wb.FullName, ".csv", ".xlsx"), FileFormat:=xlExcel8
        wb.Close True
        Set wb = Nothing
        strFile = Dir
    End If
Loop
Application.ScreenUpdating = True
End Sub
