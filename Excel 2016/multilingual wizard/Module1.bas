Attribute VB_Name = "Module1"
Public Const APPNAME As String = "Wizard Demo"

Sub StartWizard()
    UWizard.Show
End Sub

Sub del()
    
    Dim ctl As Control
    
    For Each ctl In UWizard.Controls
        On Error Resume Next
            Debug.Print ctl.Name, ctl.Caption
        On Error GoTo 0
    Next ctl
    
End Sub
