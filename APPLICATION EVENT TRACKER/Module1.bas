Attribute VB_Name = "Module1"
Option Explicit

Dim X As New clsApp
Public EventNum
 
Sub StartTrackingEvents()
    Set X.XL = Excel.Application
    EventNum = 0
    UserForm1.lblEvents.Caption = "Event Monitoring Started " & Now
    UserForm1.Show 0
End Sub

Sub StopTrackingEvents()
    Set X = Nothing
    Unload UserForm1
End Sub

