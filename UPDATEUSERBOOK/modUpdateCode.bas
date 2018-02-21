Attribute VB_Name = "modUpdateCode"
Option Explicit

Sub BeginUpdate()
    Dim Filename As String
    Dim Msg As String
    Filename = "UserBook.xlsm"
    
'   Activate workbook
    On Error Resume Next
    Workbooks(Filename).Activate
    If Err <> 0 Then
        MsgBox Filename & " must be open.", vbCritical
        Exit Sub
    End If

    Msg = "This macro will replace Module1 in UserBook.xlsm "
    Msg = Msg & "with an updated Module." & vbCrLf & vbCrLf
    Msg = Msg & "Click OK to continue."
    If MsgBox(Msg, vbInformation + vbOKCancel) = vbOK Then
        Call ReplaceModule
    Else
        MsgBox "Module not replaced,", vbCritical
    End If
End Sub

Sub ReplaceModule()
    Dim ModuleFile As String
    Dim VBP As VBIDE.VBProject

'   Export Module1 from this workbook
    ModuleFile = Application.DefaultFilePath & "\tempmodxxx.bas"
    ThisWorkbook.VBProject.VBComponents("Module1") _
      .Export ModuleFile
      
'   Replace Module1 in UserBook
    Set VBP = Workbooks("UserBook.xlsm").VBProject
    On Error GoTo ErrHandle
    With VBP.VBComponents
        .Remove VBP.VBComponents("Module1")
        .Import ModuleFile
    End With
    
'   Delete the temporary module file
    Kill ModuleFile
    MsgBox "The module has been replaced.", vbInformation
    Exit Sub

ErrHandle:
'   Did an error occur?
    MsgBox "ERROR. The module may not have been replaced.", _
      vbCritical
End Sub


