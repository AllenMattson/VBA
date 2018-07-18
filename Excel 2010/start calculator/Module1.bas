Attribute VB_Name = "Module1"
Option Explicit

Declare PtrSafe Function OpenProcess Lib "kernel32" _
    (ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, _
    lpExitCode As Long) As Long

Sub StartCalc()
    Dim Program As String
    Dim TaskID As Double
    On Error Resume Next
    Program = "calc.exe"
    TaskID = Shell(Program, 1)
    If Err <> 0 Then
        MsgBox "Cannot start " & Program, vbCritical, "Error"
    End If
End Sub

Sub StartCalc2()
    Dim TaskID As Long
    Dim hProc As Long
    Dim lExitCode As Long
    Dim ACCESS_TYPE As Integer, STILL_ACTIVE As Integer
    Dim Program As String

    ACCESS_TYPE = &H400
    STILL_ACTIVE = &H103

    Program = "calc.exe"
    On Error Resume Next

'   Shell the task
    TaskID = Shell(Program, 1)

'   Get the process handle
    hProc = OpenProcess(ACCESS_TYPE, False, TaskID)
    
    If Err <> 0 Then
        MsgBox "Cannot start " & Program, vbCritical, "Error"
        Exit Sub
    End If
    
    Do  'Loop continuously
'       Check on the process
        GetExitCodeProcess hProc, lExitCode
'       Allow event processing
        DoEvents
    Loop While lExitCode = STILL_ACTIVE
    
'   Task is finsihed, so show message
    MsgBox Program & " was closed."
End Sub


Sub ActivateCalc()
    Dim AppFile As String
    Dim CalcTaskID As Double
    
    AppFile = "Calc.exe"
    On Error Resume Next
    AppActivate "Calculator"
    If Err <> 0 Then
        Err = 0
        CalcTaskID = Shell(AppFile, 1)
        If Err <> 0 Then MsgBox "Can't start Calculator"
    End If
End Sub

