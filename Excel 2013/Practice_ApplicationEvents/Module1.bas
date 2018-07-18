Attribute VB_Name = "Module1"
Dim DoThis As New clsApplication

Public Sub InitializeAppEvents()
    Set DoThis.App = Application
End Sub

Public Sub CancelAppEvents()
    Set DoThis.App = Nothing
End Sub



