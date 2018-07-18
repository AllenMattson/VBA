Attribute VB_Name = "Module1"
Option Explicit

Sub GetCapsLockState()
    Dim CapsLock As New CapsLockClass
    MsgBox CapsLock.Value
End Sub

Sub CapsLockOn()
    Dim CapsLock As New CapsLockClass
    CapsLock.Value = True
End Sub
Sub CapsLockOff()
    Dim CapsLock As New CapsLockClass
    CapsLock.Value = False
End Sub

Sub ToggleCapsLock()
    Dim CapsLock As New CapsLockClass
    CapsLock.Toggle
End Sub


Sub GetNumLockState()
    Dim NumLock As New NumLockClass
    MsgBox NumLock.Value
End Sub

Sub NumLockOn()
    Dim NumLock As New NumLockClass
    NumLock.Value = True
End Sub
Sub NumLockOff()
    Dim NumLock As New NumLockClass
    NumLock.Value = False
End Sub
Sub ToggleNumLock()
    Dim NumLock As New NumLockClass
    NumLock.Toggle
End Sub


Sub GetScrollLockState()
    Dim ScrollLock As New ScrollLockClass
    MsgBox ScrollLock.Value
End Sub

Sub ScrollLockOn()
    Dim ScrollLock As New ScrollLockClass
    ScrollLock.Value = True
End Sub
Sub ScrollLockOff()
    Dim ScrollLock As New ScrollLockClass
    ScrollLock.Value = False
End Sub

Sub ToggleScrollLock()
    Dim ScrollLock As New ScrollLockClass
    ScrollLock.Toggle
End Sub


