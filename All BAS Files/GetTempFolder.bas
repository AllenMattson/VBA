Attribute VB_Name = "GetTempFolder"
Private Declare Function GetTempPath Lib "kernel32" Alias _
"GetTempPathA" (ByVal nBufferLength As Long, ByVal _
lpBuffer As String) As Long

Const MAX_PATH = 260

' This function uses Windows API GetTempPath to get the temporary folder

Sub Get_Temporary_Folder()

sTempFolder = GetTmpPath

End Sub


Private Function GetTmpPath()

Dim sFolder As String ' Name of the folder
Dim lRet As Long ' Return Value

sFolder = String(MAX_PATH, 0)
lRet = GetTempPath(MAX_PATH, sFolder)

If lRet <> 0 Then
GetTmpPath = Left(sFolder, InStr(sFolder, _
Chr(0)) - 1)
Else
GetTmpPath = vbNullString
End If

End Function

