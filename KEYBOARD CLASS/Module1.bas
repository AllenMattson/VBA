Attribute VB_Name = "Module1"
Option Explicit

Public Sub NumLockTest()
    Dim clsNumLock As CNumLock
    Dim OldValue As Boolean
    
    Set clsNumLock = New CNumLock
    OldValue = clsNumLock.Value
    clsNumLock.Toggle
    DoEvents 'Let the system recognize the change
    MsgBox "Num Lock was changed from " & _
        OldValue & " to " & clsNumLock.Value
End Sub

Public Sub KeyboardTest()
    Dim clsKeyboard As CKeyboard
    Dim OldValue As Boolean
    
    Set clsKeyboard = New CKeyboard
    OldValue = clsKeyboard.NumLock
    clsKeyboard.ToggleNumLock
    DoEvents
    Debug.Print "Num Lock from " & OldValue & " to " & clsKeyboard.NumLock
    
    OldValue = clsKeyboard.CapsLock
    clsKeyboard.ToggleCapsLock
    DoEvents
    Debug.Print "Caps Lock from " & OldValue & " to " & clsKeyboard.CapsLock
    
    OldValue = clsKeyboard.ScrollLock
    clsKeyboard.ToggleScrollLock
    DoEvents
    Debug.Print "Scroll Lock from " & OldValue & " to " & clsKeyboard.ScrollLock
    
End Sub
