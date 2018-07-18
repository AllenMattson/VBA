Attribute VB_Name = "Module1"
Sub ListProcedures()
    Dim VBP As VBIDE.VBProject
    Dim VBC As VBComponent
    Dim CM As CodeModule
    Dim StartLine As Long
    Dim Msg As String
    Dim ProcName As String
    
'   Use the active workbook
    Set VBP = ActiveWorkbook.VBProject
    
'   Loop through the VB components
    For Each VBC In VBP.VBComponents
        Set CM = VBC.CodeModule
        Msg = Msg & vbNewLine
        StartLine = CM.CountOfDeclarationLines + 1
        Do Until StartLine >= CM.CountOfLines
            Msg = Msg & VBC.Name & ": " & _
              CM.ProcOfLine(StartLine, vbext_pk_Proc) & vbNewLine
            StartLine = StartLine + CM.ProcCountLines _
              (CM.ProcOfLine(StartLine, vbext_pk_Proc), vbext_pk_Proc)
        Loop
    Next VBC
    MsgBox Msg
End Sub

Sub Macro1()

End Sub

Sub Macro2()

End Sub

