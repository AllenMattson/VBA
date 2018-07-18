Attribute VB_Name = "Module3"
Option Explicit

Enum SysMetConst
    x_screenWidth = SM_CXSCREEN
    y_screenHeight = SM_CYSCREEN
End Enum

Public Function ScreenRes(ByVal eIndex As SysMetConst) As Long
    ScreenRes = GetSystemMetrics(eIndex)
End Function


Sub WhatIsMyScreenResolution()
    MsgBox ScreenRes(x_screenWidth) & " x " & _
    ScreenRes(y_screenHeight)
End Sub



