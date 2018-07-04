Attribute VB_Name = "ValidateLists"
Sub ValidateLists()
Attribute ValidateLists.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
Dim Str As String, Tstr As String
Dim cell As Range
For Each cell In Sheets("Teams").Range("Conference")
cell.Offset(0, 1).Validation.Delete
    Str = cell.Value
    If InStr(cell.Value, "-") Then Str = Replace(Str, "-", "_")
    If InStr(Str, " ") Then Str = Replace(Str, " ", "_")
    With cell.Offset(0, 1).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=" + Str 'Validate List by String
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Conference"
        .ErrorTitle = ""
        .InputMessage = cell.Value
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
Next cell
End Sub
