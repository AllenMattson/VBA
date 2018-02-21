VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportSelector 
   Caption         =   "Request Report"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4905
   OleObjectBlob   =   "ReportSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReportSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdOK_Click()
    Dim Ctrl As Control
    Dim chkflag As Integer
    Dim strMsg As String
    If Me.cboDept.Value = "" Then
       MsgBox "Select the Department."
       Me.cboDept.SetFocus
       Exit Sub
    End If
    For Each Ctrl In Me.Controls
       Select Case Ctrl.Name
         Case "chk1", "chk2", "chk3"
           If Ctrl.Value = True Then
             strMsg = strMsg & vbCrLf & Ctrl.Caption

             chkflag = 1
           End If
       End Select
    Next
    If chkflag = 1 Then
      MsgBox "Run the Report(s) for " & Me.cboDept.Value & ":" & Chr(13) & Chr(13) & strMsg
    Else
      MsgBox "Please select Report type."
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Sub UserForm_Initialize()
    With Me.cboDept
        .AddItem "Marketing"
        .AddItem "Sales"
        .AddItem "Finance"
        .AddItem "Research & Development"
        .AddItem "Human Resources"
    End With
End Sub
