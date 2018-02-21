Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 And Win64 Then
    Declare PtrSafe Function GetSystemMetrics Lib "user32" _
      (ByVal nIndex As Long) As Long
#Else
Declare Function GetSystemMetrics Lib "user32" _
  (ByVal nIndex As Long) As Long
#End If


Public Const SM_CMONITORS = 80
Public Const SM_CXSCREEN = 0
Public Const SM_CYSCREEN = 1
Public Const SM_CXVIRTUALSCREEN = 78
Public Const SM_CYVIRTUALSCREEN = 79

Sub DisplayVideoInfo()
    Dim numMonitors As Long
    Dim vidWidth As Long, vidHeight As Long
    Dim virtWidth As Long, virtHeight As Long
    Dim Msg As String
    
    numMonitors = GetSystemMetrics(SM_CMONITORS)
    vidWidth = GetSystemMetrics(SM_CXSCREEN)
    vidHeight = GetSystemMetrics(SM_CYSCREEN)
    virtWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN)
    virtHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN)

    If numMonitors > 1 Then
        Msg = numMonitors & " display monitors" & vbCrLf
        Msg = Msg & "Virtual screen: " & virtWidth & " X "
        Msg = Msg & virtHeight & vbCrLf & vbCrLf
        Msg = Msg & "The video mode on the primary display is: "
        Msg = Msg & vidWidth & " X " & vidHeight
    Else
        Msg = Msg & "The video display mode: "
        Msg = Msg & vidWidth & " X " & vidHeight
    End If
    MsgBox Msg
End Sub



