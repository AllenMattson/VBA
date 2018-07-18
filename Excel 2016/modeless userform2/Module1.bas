Attribute VB_Name = "Module1"
Option Explicit

Public Const APPNAME As String = "Power Programming Cell Info"
Public gclsEvent As CEvent
Public CheckingCells As Boolean
' CheckingCells is set to True in the procedures that check for
' cell dependents and precedents. This is because the Precedents,
' DirectPrecedents, Dependents, and DirectDependents methods trigger
' a SelectionChange event. The SelectionChange event handler in Class1
' checks the value of CheckingCells and exits if it's True.

Sub ShowCellInfoBox()
    Dim Ans As Long
    
    If Application.Version < 9 Then
        MsgBox "Excel 2000 or later is required.", vbCritical
        Exit Sub
    End If
    If ActiveSheet.ProtectContents Then
        Ans = MsgBox("The active sheet is protected. This utility will still work, but it will not display information about cell dependents and precedents." & vbCrLf & vbCrLf & "Continue?", vbQuestion + vbYesNo, APPNAME)
        If Ans <> vbYes Then
            Exit Sub
        End If
    End If
    Set gclsEvent = New CEvent
    Set gclsEvent.app = Application
    CheckingCells = False
    UserForm1.Show vbModeless
End Sub

