Attribute VB_Name = "Close_ImmediateWin"
Option Explicit

Sub Close_ImmediateWin()
    Dim objWin As VBIDE.Window
    Dim strOpenWindows As String

    strOpenWindows = "The following windows are open:" & _
           vbCrLf & vbCrLf

    For Each objWin In Application.VBE.Windows
        Select Case objWin.Type
            Case vbext_wt_Immediate
                MsgBox objWin.Caption & " window was closed."
                objWin.Close
            Case Else
                strOpenWindows = strOpenWindows & _
                    objWin.Caption & vbCrLf
        End Select
    Next
    MsgBox strOpenWindows
    Set objWin = Nothing
End Sub



