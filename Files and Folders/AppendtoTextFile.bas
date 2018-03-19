Attribute VB_Name = "AppendtoTextFile"
Sub VBA_to_append_existing_text_file(Str As String)
Dim strFile_Path As String: strFile_Path = "C:\temp\test.txt" 'Change as per your test folder and exiting file path to append it.
Open strFile_Path For Append As #1
Write #1, Str
Close #1
End Sub
