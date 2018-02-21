Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Sub getMouseCoordinates()
    Dim cPos As POINTAPI

    GetCursorPos cPos
    Debug.Print "x coordinate:" & cPos.x
    Debug.Print "y coordinate:" & cPos.y
End Sub


