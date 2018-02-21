Attribute VB_Name = "Module1"
Option Explicit

Const TOOLBARNAME As String = "MyToolbar"

Sub CreateToolbar()
    Dim TBar As CommandBar
    Dim Btn As CommandBarButton
    
'   Delete existing toolbar (if it exists)
    On Error Resume Next
    CommandBars(TOOLBARNAME).Delete
    On Error GoTo 0
    
'   Create toolbar
    Set TBar = CommandBars.Add
    With TBar
        .Name = TOOLBARNAME
        .Visible = True
    End With
    
'   Add a button
    Set Btn = TBar.Controls.Add(Type:=msoControlButton)
    With Btn
       .FaceId = 300
       '.Picture = Application.CommandBars.GetImageMso("ViewAppointmentInCalendar", 16, 16)
       .OnAction = "Macro1"
       .Caption = "Macro1 Tooltip goes here"
    End With

'   Add another button
    Set Btn = TBar.Controls.Add(Type:=msoControlButton)
    With Btn
       .FaceId = 25
       '.Picture = Application.CommandBars.GetImageMso("CDAudioStopTime", 16, 16)
       .OnAction = "Macro2"
       .Caption = "Macro2 Tooltip goes here"
    End With

End Sub

Sub DeleteToolbar()
    On Error Resume Next
    CommandBars(TOOLBARNAME).Delete
    On Error GoTo 0
End Sub


Sub Macro1()
    MsgBox "Today is " & Date, vbInformation
End Sub

Sub Macro2()
    MsgBox "Currently, it's " & Time, vbInformation
End Sub

