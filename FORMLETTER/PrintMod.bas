Attribute VB_Name = "PrintMod"
Sub PrintForms()
Attribute PrintForms.VB_ProcData.VB_Invoke_Func = " \n14"
    Sheets("Form").Activate
    StartRow = Range("StartRow")
    EndRow = Range("EndRow")
    
    If StartRow > EndRow Then
        Msg = "ERROR" & Chr(13) & "The starting row must be less than the ending row!"
        MsgBox Msg, vbCritical, APPNAME
    End If
    
    For i = StartRow To EndRow
        Range("RowIndex") = i
        ActiveSheet.PrintOut Copies:=1
    Next i
End Sub

Sub EditData()
Attribute EditData.VB_ProcData.VB_Invoke_Func = " \n14"
    Worksheets("Data").Activate
    Range("A1").Select
End Sub

Sub ReturnToForm()
Attribute ReturnToForm.VB_ProcData.VB_Invoke_Func = " \n14"
    Worksheets("Form").Activate
    Range("RowIndex").Select
End Sub
