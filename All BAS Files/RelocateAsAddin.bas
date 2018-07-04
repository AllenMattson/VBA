Attribute VB_Name = "RelocateAsAddin"
Sub AddN()
Dim TEST1 As String
Dim TEST2 As String
Dim TEST3 As String
Workbooks.Open Filename:="C:\Users\NewFile\AddInTest.xlsm"
TEST1 = "C:\Users\"
TEST2 = "AppData\Roaming\Microsoft\AddIns\AddInTest2.xla"
TEST3 = Environ("username") & "\"
ActiveWorkbook.IsAddin = True
Application.DisplayAlerts = False
ActiveWorkbook.SaveAs Filename:=TEST1 & TEST3 & TEST2, FileFormat:=xlOpenXMLAddIn
Application.DisplayAlerts = True
ChDir "C:\Users\" & TEST3 & "\AppData\Roaming\Microsoft\AddIns"
Windows("AddInTest").Close
End Sub

