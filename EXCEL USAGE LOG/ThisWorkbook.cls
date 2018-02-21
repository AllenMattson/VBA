VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Open Application.DefaultFilePath & "\excelusage.txt" For Append As #1
    Print #1, "Started " & Now
    Close #1
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Open Application.DefaultFilePath & "\excelusage.txt" _
      For Append As #1
    Print #1, "Stopped " & Now
    Close #1
End Sub


