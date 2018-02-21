Attribute VB_Name = "ShowSurvey"
Option Explicit

Sub DoSurvey()
    InfoSurvey.Show
End Sub


Sub ListHardware()
    With InfoSurvey.lboxSystems
        .AddItem "DVD Drive"
        .AddItem "Printer"
        .AddItem "Fax"
        .AddItem "Network"
        .AddItem "Joystick"
        .AddItem "Sound Card"
        .AddItem "Graphics Card"
        .AddItem "Modem"
        .AddItem "Monitor"
        .AddItem "Mouse"
        .AddItem "External Drive"
        .AddItem "Scanner"
    End With
End Sub

Sub ListSoftware()
    With InfoSurvey.lboxSystems
        .AddItem "Spreadsheets"
        .AddItem "Databases"
        .AddItem "CAD Systems"
        .AddItem "Word Processing"
        .AddItem "Finance Programs"
        .AddItem "Games"
        .AddItem "Accounting Programs"
        .AddItem "Desktop Publishing"
        .AddItem "Imaging Software"
        .AddItem "Personal Information Managers"
    End With
End Sub
