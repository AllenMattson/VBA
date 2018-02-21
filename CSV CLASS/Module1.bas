Attribute VB_Name = "Module1"
Option Explicit

Sub ExportARange()
    Dim RangeToExport As Range
    
    Set RangeToExport = ActiveWindow.RangeSelection
    If Application.CountA(RangeToExport) = 0 Then
        MsgBox "The selection is empty."
        Exit Sub
    End If
    
    Dim CSVFile As New CSVFileClass
    On Error Resume Next
    With CSVFile
        .ExportRange = RangeToExport
        .Export CSVFileName:=Application.DefaultFilePath & "\temp.csv"
        If Err <> 0 Then MsgBox "Cannot export" & Application.DefaultFilePath & "\temp.csv"
    End With
End Sub

Sub ImportAFile()
    Dim CSVFile As New CSVFileClass
    On Error Resume Next
    With CSVFile
        .ImportRange = ActiveCell
        .Import CSVFileName:=Application.DefaultFilePath & "\temp.csv"
        If Err <> 0 Then MsgBox "Cannot import " & Application.DefaultFilePath & "\temp.csv"
    End With
End Sub

