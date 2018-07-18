VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Select A VB Project"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub OKButton_Click()
    Call ShowComponents(ListBox1.Value)
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim VBP As VBIDE.VBProject
    Dim ProtectedProjects As Long
    Dim FileName As String
    
    For Each VBP In Application.VBE.VBProjects
        If VBP.Protection = 1 Then
            ProtectedProjects = ProtectedProjects + 1
        Else
            On Error Resume Next
            FileName = FileNameOnly(VBP.FileName)
            If Err.Number = 0 And FileName <> "" Then
                ListBox1.AddItem FileName
            Else
                ProtectedProjects = ProtectedProjects + 1
                On Error GoTo 0
            End If
        End If
    Next VBP
    If ProtectedProjects > 0 Then MsgBox ProtectedProjects & " protected VB projects are not listed."
End Sub

Private Function FileNameOnly(FilePath)
    Dim x As Variant
    On Error Resume Next
    FileNameOnly = ""
    x = Split(FilePath, "\")
    FileNameOnly = x(UBound(x))
End Function

Private Sub CancelButton_Click()
    Unload Me
End Sub

