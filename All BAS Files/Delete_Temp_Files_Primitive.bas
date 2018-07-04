Attribute VB_Name = "Delete_Temp_Files_Primitive"
Sub Delete_Temp_Files_Primitive()

Dim sFileType As String ' Declare the Type of File
Dim sTempDir As String ' Temporary Directory

' ---------------------------------------------------------------
' Written By Shanmuga Sundara Raman for http://vbadud.blogspot.com
' ---------------------------------------------------------------

On Error Resume Next

sFileType = "*.tmp"
sTempDir = "c:\windows\Temp\" ' There might be mutiple temp directories (one for each profile) in Windows XP. Modify the code accordingly

Kill sTempDir & sFileType

' ---------------------------------------------------------------
' Delete Temporary Files, Excel VBA, Kill Statement
' ---------------------------------------------------------------

End Sub
