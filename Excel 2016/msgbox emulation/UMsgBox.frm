VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UMsgBox 
   Caption         =   "Microsoft Excel"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   OleObjectBlob   =   "UMsgBox.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public UserClick As Long

'RETURN VALUES CONSTANTS
'vbOK = 1 OK
'vbCancel = 2 Cancel
'vbAbort = 3   Abort
'vbRetry = 4   Retry
'vbIgnore = 5   Ignore
'vbYes = 6 Yes
'vbNo = 7 No

Private Sub cmdLeft_Click()
    ProcessButton Me.cmdLeft.Caption
    Me.Hide
End Sub

Private Sub cmdMiddle_Click()
    ProcessButton Me.cmdMiddle.Caption
    Me.Hide
End Sub

Private Sub cmdRight_Click()
    ProcessButton Me.cmdRight.Caption
    Me.Hide
End Sub

Private Sub ProcessButton(sCaption As String)
    Select Case sCaption
        Case "OK": Me.UserClick = vbOK
        Case "Cancel": Me.UserClick = vbCancel
        Case "Abort": Me.UserClick = vbAbort
        Case "Retry": Me.UserClick = vbRetry
        Case "Ignore": Me.UserClick = vbIgnore
        Case "Yes": Me.UserClick = vbYes
        Case "No": Me.UserClick = vbNo
    End Select
End Sub
