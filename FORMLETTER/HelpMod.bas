Attribute VB_Name = "HelpMod"
Public Const APPNAME As String = "Elephants 4U"

Sub ShowHelp()
Attribute ShowHelp.VB_ProcData.VB_Invoke_Func = " \n14"
    Msg = "Help information is stored in cell comments. Move the mouse pointer over a cell that contains a small red comment indicator. Help for the cell with appear automatically."
    MsgBox Msg, vbInformation, APPNAME
End Sub
