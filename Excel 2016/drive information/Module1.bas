Attribute VB_Name = "Module1"
#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetDriveType32 Lib "kernel32" _
        Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
    
    Private Declare PtrSafe Function GetLogicalDriveStrings Lib "kernel32" _
      Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
      ByVal lpBuffer As String) As Long
    
    Private Declare PtrSafe Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" _
        (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, _
        lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
#Else
    Private Declare Function GetDriveType32 Lib "kernel32" _
        Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
    
    Private Declare Function GetLogicalDriveStrings Lib "kernel32" _
      Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
      ByVal lpBuffer As String) As Long
    
    Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" _
        (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, _
        lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
#End If

Function DriveType(DriveLetter As String) As String
'   Returns a string that describes the drive type
    DLetter = Left(DriveLetter, 1) & ":"
    DriveCode = GetDriveType32(DLetter)
    Select Case DriveCode
        Case 1: DriveType = "Local"
        Case 2: DriveType = "Removable"
        Case 3: DriveType = "Fixed"
        Case 4: DriveType = "Remote"
        Case 5: DriveType = "CD-ROM"
        Case 6: DriveType = "RAM Disk"
        Case Else: DriveType = "Unknown"
    End Select
End Function


Function NumberofDrives() As Integer
'   Returns the number of drives
    Dim Buffer As String * 255
    Dim BuffLen As Long
    Dim DriveCount As Integer
   
    BuffLen = GetLogicalDriveStrings(Len(Buffer), Buffer)
    DriveCount = 0
'   Search for a null -- which separates the drives
    For i = 1 To BuffLen
        If Asc(Mid(Buffer, i, 1)) = 0 Then _
          DriveCount = DriveCount + 1
    Next i
    NumberofDrives = DriveCount
End Function


Function DriveName(index As Integer) As String
'   Returns the drive letter using an index
'   Returns an empty string if index > number of drives
    
    Dim Buffer As String * 255
    Dim BuffLen As Long
    Dim TheDrive As String
    Dim DriveCount As Integer
   
    BuffLen = GetLogicalDriveStrings(Len(Buffer), Buffer)

'   Search thru the string of drive names
    TheDrive = ""
    DriveCount = 0
    For i = 1 To BuffLen
        If Asc(Mid(Buffer, i, 1)) <> 0 Then _
          TheDrive = TheDrive & Mid(Buffer, i, 1)
        If Asc(Mid(Buffer, i, 1)) = 0 Then 'null separates drives
            DriveCount = DriveCount + 1
            If DriveCount = index Then
                DriveName = UCase(Left(TheDrive, 1))
                Exit Function
            End If
            TheDrive = ""
        End If
    Next i
End Function


Sub ShowDriveInfo()
'   This sub writes information for all drives
'   to a range of cells
'   Demonstrates the use of the custom drive functions

    Dim cuAvailable As Currency
    Dim cuTotal As Currency
    Dim cuFree As Currency
    
    Dim i As Integer
    Dim DLetter As String
    Dim NumDrives As Integer
        
    NumDrives = NumberofDrives()

'   Write info for all drives to active cell location
    Cells.ClearContents
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell"
        Exit Sub
    End If
    
'   Insert headings
    Application.ScreenUpdating = False
    With ActiveCell
        .Offset(0, 0).Value = "Drive"
        .Offset(0, 1).Value = "Type"
        .Offset(0, 2).Value = "Total Bytes"
        .Offset(0, 3).Value = "Used Bytes"
        .Offset(0, 4).Value = "Free Bytes"
        
'   Insert data for each drive
    For i = 1 To NumDrives
        DLetter = DriveName(i) & ":\"
        cuAvailable = 0
        cuTotal = 0
        cuFree = 0
        Call GetDiskFreeSpaceEx(DLetter, cuAvailable, cuTotal, cuFree)
    
'       Drive name
        .Offset(i, 0).Value = DLetter
'       Drive type
        .Offset(i, 1) = DriveType(DLetter)
'       Total space
        .Offset(i, 2) = Format(cuTotal * 10000, "#,###")
'       Used space
        .Offset(i, 3) = Format((cuTotal - cuFree) * 10000, "#,###")
'       Free space
        .Offset(i, 4) = Format(cuFree * 10000, "#,###")

    Next i
'   Format the table
    ActiveSheet.ListObjects.Add xlSrcRange, ActiveCell.CurrentRegion
    With ActiveCell.ListObject
        .TableStyle = "TableStyleLight8"
        .ShowTableStyleRowStripes = False
        .ShowTableStyleColumnStripes = True
    End With
    .CurrentRegion.Columns.AutoFit
    End With 'ActiveCell
End Sub




