Attribute VB_Name = "ExtractEmailFromStrings"
Sub ExtractEmail()
'Update 20130829
Dim WorkRng As Range
Dim arr As Variant
Dim CharList As String
On Error Resume Next
xTitleId = "KutoolsforExcel"
Set WorkRng = Application.Selection
Set WorkRng = Application.InputBox("Range", xTitleId, WorkRng.Address, Type:=8)
arr = WorkRng.Value
CheckStr = "[A-Za-z0-9._-]"
For i = 1 To UBound(arr, 1)
    For j = 1 To UBound(arr, 2)
        extractStr = arr(i, j)
        outStr = ""
        Index = 1
        Do While True
            Index1 = VBA.InStr(Index, extractStr, "@")
            getStr = ""
            If Index1 > 0 Then
                For p = Index1 - 1 To 1 Step -1
                    If Mid(extractStr, p, 1) Like CheckStr Then
                        getStr = Mid(extractStr, p, 1) & getStr
                    Else
                        Exit For
                    End If
                Next
                getStr = getStr & "@"
                For p = Index1 + 1 To Len(extractStr)
                    If Mid(extractStr, p, 1) Like CheckStr Then
                        getStr = getStr & Mid(extractStr, p, 1)
                    Else
                        Exit For
                    End If
                Next
                Index = Index1 + 1
                If outStr = "" Then
                    outStr = getStr
                Else
                    outStr = outStr & Chr(10) & getStr
                End If
            Else
                Exit Do
            End If
        Loop
        arr(i, j) = outStr
    Next
Next
WorkRng.Value = arr
End Sub
