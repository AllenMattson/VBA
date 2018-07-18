Attribute VB_Name = "Module1"
Option Explicit

Sub ShowDriveInfo()
    Dim FileSys As FileSystemObject
    Dim Drv As Drive
    Dim Row As Long
    Set FileSys = CreateObject("Scripting.FileSystemObject")
    Cells.ClearContents
    Row = 1
'   Column headers
    Range("A1:F1") = Array("Drive", "Ready", "Type", "Vol. Name", _
      "Size", "Available")
    On Error Resume Next
'   Loop through the drives
    For Each Drv In FileSys.Drives
        Row = Row + 1
        Cells(Row, 1) = Drv.DriveLetter
        Cells(Row, 2) = Drv.IsReady
        Select Case Drv.DriveType
            Case 0: Cells(Row, 3) = "Unknown"
            Case 1: Cells(Row, 3) = "Removable"
            Case 2: Cells(Row, 3) = "Fixed"
            Case 3: Cells(Row, 3) = "Network"
            Case 4: Cells(Row, 3) = "CD-ROM"
            Case 5: Cells(Row, 3) = "RAM Disk"
        End Select
        Cells(Row, 4) = Drv.VolumeName
        Cells(Row, 5) = Drv.TotalSize
        Cells(Row, 6) = Drv.AvailableSpace
    Next Drv
    'Make a table
    ActiveSheet.ListObjects.Add xlSrcRange, _
      Range("A1").CurrentRegion, , xlYes
End Sub

