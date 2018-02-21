Attribute VB_Name = "Module1"
Option Explicit

Sub ExportRange()
    Dim Filename As String
    Dim NumRows As Long, NumCols As Integer
    Dim r As Long, c As Integer
    Dim Data
    Dim ExpRng As Range
    Set ExpRng = Application.Intersect(Selection, ActiveSheet.UsedRange)
    NumCols = ExpRng.Columns.Count
    NumRows = ExpRng.Rows.Count
    Filename = Application.DefaultFilePath & "\textfile.csv"
    Open Filename For Output As #1
        For r = 1 To NumRows
            For c = 1 To NumCols
                Data = ExpRng.Cells(r, c).Value
                If IsNumeric(Data) Then Data = Val(Data)
                If IsEmpty(ExpRng.Cells(r, c)) Then Data = ""
                If c <> NumCols Then
                    Write #1, Data;
                Else
                    Write #1, Data
                End If
            Next c
        Next r
    Close #1
    MsgBox ExpRng.Count & " cells were exported to " & Filename, vbInformation
End Sub



Sub ImportRange()
    Dim ImpRng As Range
    Dim Filename As String
    Dim r As Long, c As Integer
    Dim txt As String, Char As String * 1
    Dim Data
    Dim i As Integer
    
    Set ImpRng = ActiveCell
    On Error Resume Next
    Filename = Application.DefaultFilePath & "\textfile.csv"
    Open Filename For Input As #1
    If Err <> 0 Then
        MsgBox "Not found: " & Filename, vbCritical, "ERROR"
        Exit Sub
    End If
    r = 0
    c = 0
    txt = ""
    Application.ScreenUpdating = False
    Do Until EOF(1)
        Line Input #1, Data
        For i = 1 To Len(Data)
            Char = Mid(Data, i, 1)
            If Char = "," Then 'comma
                ActiveCell.Offset(r, c) = txt
                c = c + 1
                txt = ""
            ElseIf i = Len(Data) Then 'end of line
                If Char <> Chr(34) Then txt = txt & Char
                ActiveCell.Offset(r, c) = txt
                txt = ""
            ElseIf Char <> Chr(34) Then
                txt = txt & Char
            End If
        Next i
        c = 0
        r = r + 1
    Loop
    Close #1
    Application.ScreenUpdating = True
End Sub


Function DefaultFilePath()
    DefaultFilePath = Application.DefaultFilePath
End Function
