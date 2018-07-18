Attribute VB_Name = "Module1"
Option Explicit

Sub test()
Dim p As Picture
p.ShapeRange.PictureFormat.TransparentBackground
End Sub

Sub ShowFaceIDs()
    Dim NewToolbar As CommandBar
    Dim ctl As CommandBarButton
    Dim ID_Start As Integer, ID_End As Integer
    Dim TopPos As Long, LeftPos As Long
    Dim i As Long, Count As Long

    On Error Resume Next
    ID_Start = Range("FirstID").Value
    ID_End = Range("LastID").Value
    If Err.Number <> 0 Or (ID_Start > ID_End) Then
        MsgBox "Error - check the ID values", vbCritical
        Exit Sub
    End If

'   Delete existing FaceIds toolbar if it exists
    On Error Resume Next
    Application.CommandBars("TempFaceIds").Delete
    On Error GoTo 0

'   Clear the sheet
    ActiveSheet.Pictures.Delete
    Application.ScreenUpdating = False
    
'   Add an empty toolbar
    Set NewToolbar = Application.CommandBars.Add _
        (Name:="TempFaceIds", temporary:=True)
    NewToolbar.Visible = True

    TopPos = 60
    LeftPos = 16
    Count = 1
    For i = ID_Start To ID_End
        On Error Resume Next
        NewToolbar.Controls(1).Delete
        On Error GoTo 0
        Set ctl = NewToolbar.Controls.Add(Type:=msoControlButton)
        ctl.FaceId = i
        ctl.CopyFace
        ActiveSheet.Paste
        'On Error Resume Next
        With ActiveSheet.Pictures(Count)
            .Top = TopPos
            .Left = LeftPos
            .Name = "FaceID " & i
        End With
        LeftPos = LeftPos + 16
        If Count Mod 40 = 0 Then
            TopPos = TopPos + 16
            LeftPos = 16
        End If
        Count = Count + 1
    Next i
    ActiveWindow.RangeSelection.Select
'   Delete toolbar
    Application.CommandBars("TempFaceIds").Delete
End Sub
