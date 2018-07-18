Attribute VB_Name = "Module3"
Option Explicit

Sub Discover_EmailSystem()
    Select Case Application.MailSystem
        Case xlMAPI
            MsgBox "You have Microsoft Mail installed."
        Case xlNoMailSystem
            MsgBox "No mail system installed on this computer."
        Case xlPowerTalk
            MsgBox "Your mail system is PowerTalk"
    End Select
End Sub


