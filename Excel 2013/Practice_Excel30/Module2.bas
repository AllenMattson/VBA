Attribute VB_Name = "Module2"
Public Const SM_CXSCREEN = 0  'defines the screen width
    Public Const SM_CYSCREEN = 1  'defines the screen height

Public Declare Function GetSystemMetrics Lib "user32" _
    (ByVal nIndex As Long) As Long

Sub GetScreenResolution()
    Dim xSM As Long
    Dim ySM As Long

    xSM = GetSystemMetrics(SM_CXSCREEN)
    ySM = GetSystemMetrics(SM_CYSCREEN)

    Debug.Print "Your screen resolution is: " & _
        xSM & " x " & ySM
End Sub



