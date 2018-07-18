Attribute VB_Name = "Module1"
Option Explicit

Sub AddButtonAndCode()
Attribute AddButtonAndCode.VB_Description = "Macro recorded 12/17/1998 by John Walkenbach"
Attribute AddButtonAndCode.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim NewSheet As Worksheet
    Dim NewButton As OLEObject
    Dim Code As String
    Dim NextLine As Long
    Dim x
    
'   Make sure access to the VBProject is allowed
    On Error Resume Next
    Set x = ActiveWorkbook.VBProject
    If Err <> 0 Then
        MsgBox "Your security settings do not allow this macro to run.", vbCritical
        On Error GoTo 0
        Exit Sub
    End If

'   Add the sheet
    Set NewSheet = Sheets.Add
    
'   Add a CommandButton
    Set NewButton = NewSheet.OLEObjects.Add _
      ("Forms.CommandButton.1")
    With NewButton
        .Left = 4
        .Top = 4
        .Width = 150
        .Height = 36
        .Object.Caption = "Return to Sheet1"
    End With
    
'   Add the event handler code
    Code = "Sub CommandButton1_Click()" & vbNewLine
    Code = Code & "    On Error Resume Next" & vbNewLine
    Code = Code & "    Sheets(""Sheet1"").Activate" & vbNewLine
    Code = Code & "    If Err <> 0 Then" & vbNewLine
    Code = Code & "      MsgBox ""Cannot activate Sheet1.""" & vbNewLine
    Code = Code & "    End If" & vbNewLine
    Code = Code & "End Sub"
    
    With ActiveWorkbook.VBProject. _
      VBComponents(NewSheet.Name).CodeModule
        NextLine = .CountOfLines + 1
        .InsertLines NextLine, Code
    End With
End Sub
