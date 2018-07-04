Attribute VB_Name = "Module2"
Sub unprotected()
    If Hook Then
        MsgBox "VBA Project is unprotected!", vbInformation, "*****"
    End If
End Sub
