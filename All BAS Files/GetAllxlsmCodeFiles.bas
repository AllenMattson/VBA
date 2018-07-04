Attribute VB_Name = "GetAllxlsmCodeFiles"
Sub InitialLoop()

Dim SH As Worksheet: Set SH = ThisWorkbook.Sheets("Sheet1")
Dim SH2 As Worksheet: Set SH2 = ThisWorkbook.Sheets("Sheet2")
Dim WB2 As Workbook
Dim WB As Workbook: Set WB = ThisWorkbook
Dim WKBKstr As String, ModName As String

With Application
    .DisplayAlerts = False
    .Calculation = xlCalculationManual
    .ScreenUpdating = False
End With

SH.Activate
Dim LR As Long, i As Long, j As Long
Dim TheLines As String
With SH
LR = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 739 To LR
        WKBKstr = SH.Cells(i, 1).Value
        SH2.Activate
        With SH2
            Range("A90000").End(xlUp).Offset(1, 0).Value = WKBKstr
            Set WB2 = Workbooks.Open(WKBKstr)
            With WB2
                On Error Resume Next
                For j = 1 To WB2.VBProject.VBComponents.Count
                    ModName = WB2.VBProject.VBComponents(j).Name
                    TheLines = WB2.VBProject.VBComponents(j).CodeModule.CountOfLines
                    If WB2.VBProject.VBComponents(j).CodeModule.CountOfLines > 1 Then
                        Debug.Print WB2.Name & vbTab & ModName
                        WB.Activate: SH2.Activate
                        SH2.Range("A90000").End(xlUp).Offset(1, 1).Value = TheLines
                        SH2.Range("A90000").End(xlUp).Offset(1, 0).Value = ModName
                        'DoEvents
                        WB2.Activate
                    End If
                Next j
            End With
            If ActiveWorkbook.Name = WB2.Name Then WB2.Close False
            WB.Activate
            WB.Save
        End With
            SH.Activate
    Next i
End With

With Application
    .DisplayAlerts = True
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
End With
End Sub
Sub Modules_namen()
For j = 1 To ThisWorkbook.VBProject.VBComponents.Count
    MsgBox ThisWorkbook.VBProject.VBComponents(j).Name
Next
End Sub
