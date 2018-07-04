Attribute VB_Name = "Module1"
Public oWMISrvEx As Object 'SWbemServicesEx
Public oWMIObjSet As Object 'SWbemServicesObjectSet
Public oWMIObjEx As Object 'SWbemObjectEx
Public oWMIProp As Object 'SWbemProperty
Public sWQL As String 'WQL Statement
Public n As Long
Public strRow As String
Public intRow As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'The special “Win32_” objects you need to use to access this information about your computer are as follows:
'    Win32_NetworkAdapterConfiguration – All of your network configuration settings
'    Win32_LogicalDisk – Disks with capacities and free space.
'    Win32_Processor – CPU Specs
'    Win32_PhysicalMemoryArray – RAM/Installed Memory size
'    Win32_VideoController – Graphics adapter and settings
'    Win32_OnBoardDevice – Motherboard devices
'    Win32_OperatingSystem – Which version of Windows with Serial Number
'    WIn32_Printer – Installed Printers
'    Win32_Product – Installed Software
'    Win32_BaseService – List services running (or stopped) on any PC along with the service’s path and file name.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub NetworkWMI()

sWQL = "Select * From Win32_NetworkAdapterConfiguration"
Set oWMISrvEx = GetObject("winmgmts:root/CIMV2")
Set oWMIObjSet = oWMISrvEx.ExecQuery(sWQL)
intRow = 2
strRow = Str(intRow)

ThisWorkbook.Sheets("Network").Range("A1").Value = "Name"
ThisWorkbook.Sheets("Network").Cells(1, 1).Font.Bold = True

ThisWorkbook.Sheets("Network").Range("B1").Value = "Value"
ThisWorkbook.Sheets("Network").Cells(1, 2).Font.Bold = True

For Each oWMIObjEx In oWMIObjSet

For Each oWMIProp In oWMIObjEx.Properties_
If Not IsNull(oWMIProp.Value) Then
If IsArray(oWMIProp.Value) Then
For n = LBound(oWMIProp.Value) To UBound(oWMIProp.Value)
Debug.Print oWMIProp.Name & "(" & n & ")", oWMIProp.Value(n)
ThisWorkbook.Sheets("Network").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).Value = oWMIProp.Value(n)
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
Next
Else
ThisWorkbook.Sheets("Network").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).Value = oWMIProp.Value
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
End If
End If
Next
'End If
Next
End Sub
'''***************************************************************************************************
Sub LogicalDiskWMI()

sWQL = "Select * From Win32_LogicalDisk"
Set oWMISrvEx = GetObject("winmgmts:root/CIMV2")
Set oWMIObjSet = oWMISrvEx.ExecQuery(sWQL)
intRow = 2
strRow = Str(intRow)

ThisWorkbook.Sheets("LogicalDisk").Range("A1").Value = "Name"
ThisWorkbook.Sheets("LogicalDisk").Cells(1, 1).Font.Bold = True

ThisWorkbook.Sheets("LogicalDisk").Range("B1").Value = "Value"
ThisWorkbook.Sheets("LogicalDisk").Cells(1, 2).Font.Bold = True

For Each oWMIObjEx In oWMIObjSet

For Each oWMIProp In oWMIObjEx.Properties_
If Not IsNull(oWMIProp.Value) Then
If IsArray(oWMIProp.Value) Then
For n = LBound(oWMIProp.Value) To UBound(oWMIProp.Value)
Debug.Print oWMIProp.Name & "(" & n & ")", oWMIProp.Value(n)
ThisWorkbook.Sheets("LogicalDisk").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("LogicalDisk").Range("B" & Trim(strRow)).Value = oWMIProp.Value(n)
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
Next
Else
ThisWorkbook.Sheets("Network").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).Value = oWMIProp.Value
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
End If
End If
Next
'End If
Next
End Sub



Sub ProcessorWMI()

sWQL = "Select * From win32_processor"  'CPU specs
Set oWMISrvEx = GetObject("winmgmts:root/CIMV2")
Set oWMIObjSet = oWMISrvEx.ExecQuery(sWQL)
intRow = 2
strRow = Str(intRow)

ThisWorkbook.Sheets("Processor").Range("A1").Value = "Name"
ThisWorkbook.Sheets("Processor").Cells(1, 1).Font.Bold = True

ThisWorkbook.Sheets("Processor").Range("B1").Value = "Value"
ThisWorkbook.Sheets("Processor").Cells(1, 2).Font.Bold = True

For Each oWMIObjEx In oWMIObjSet

For Each oWMIProp In oWMIObjEx.Properties_
If Not IsNull(oWMIProp.Value) Then
If IsArray(oWMIProp.Value) Then
For n = LBound(oWMIProp.Value) To UBound(oWMIProp.Value)
Debug.Print oWMIProp.Name & "(" & n & ")", oWMIProp.Value(n)
ThisWorkbook.Sheets("Processor").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).Value = oWMIProp.Value(n)
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
Next
Else
ThisWorkbook.Sheets("Processor").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Processor").Range("B" & Trim(strRow)).Value = oWMIProp.Value
ThisWorkbook.Sheets("Processor").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
End If
End If
Next
'End If
Next
End Sub
''******************************************************************************************************************************
Sub PhysicalMemWMI()

sWQL = "Select * From Win32_PhysicalMemoryArray"    'RAM/Installed Memory Size
Set oWMISrvEx = GetObject("winmgmts:root/CIMV2")
Set oWMIObjSet = oWMISrvEx.ExecQuery(sWQL)
intRow = 2
strRow = Str(intRow)

ThisWorkbook.Sheets("Physical Memory").Range("A1").Value = "Name"
ThisWorkbook.Sheets("Physical Memory").Cells(1, 1).Font.Bold = True

ThisWorkbook.Sheets("Physical Memory").Range("B1").Value = "Value"
ThisWorkbook.Sheets("Physical Memory").Cells(1, 2).Font.Bold = True

For Each oWMIObjEx In oWMIObjSet

For Each oWMIProp In oWMIObjEx.Properties_
If Not IsNull(oWMIProp.Value) Then
If IsArray(oWMIProp.Value) Then
For n = LBound(oWMIProp.Value) To UBound(oWMIProp.Value)
Debug.Print oWMIProp.Name & "(" & n & ")", oWMIProp.Value(n)
ThisWorkbook.Sheets("Physical Memory").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).Value = oWMIProp.Value(n)
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
Next
Else
ThisWorkbook.Sheets("Physical Memory").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Physical Memory").Range("B" & Trim(strRow)).Value = oWMIProp.Value
ThisWorkbook.Sheets("Physical Memory").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
End If
End If
Next
'End If
Next
End Sub
''******************************************************************************************************************************
Sub VideoControlWMI()

sWQL = "Select * From Win32_VideoController"    'GRAPHICS ADAPTER SETTINGS
Set oWMISrvEx = GetObject("winmgmts:root/CIMV2")
Set oWMIObjSet = oWMISrvEx.ExecQuery(sWQL)
intRow = 2
strRow = Str(intRow)

ThisWorkbook.Sheets("Video Controller").Range("A1").Value = "Name"
ThisWorkbook.Sheets("Video Controller").Cells(1, 1).Font.Bold = True

ThisWorkbook.Sheets("Video Controller").Range("B1").Value = "Value"
ThisWorkbook.Sheets("Video Controller").Cells(1, 2).Font.Bold = True

For Each oWMIObjEx In oWMIObjSet

For Each oWMIProp In oWMIObjEx.Properties_
If Not IsNull(oWMIProp.Value) Then
If IsArray(oWMIProp.Value) Then
For n = LBound(oWMIProp.Value) To UBound(oWMIProp.Value)
Debug.Print oWMIProp.Name & "(" & n & ")", oWMIProp.Value(n)
ThisWorkbook.Sheets("Video Controller").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).Value = oWMIProp.Value(n)
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
Next
Else
ThisWorkbook.Sheets("Video Controller").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Video Controller").Range("B" & Trim(strRow)).Value = oWMIProp.Value
ThisWorkbook.Sheets("Video Controller").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
End If
End If
Next
'End If
Next
End Sub
''******************************************************************************************************************************
Sub OnBoardWMI()

sWQL = "Select * From Win32_OnBoardDevice"  'Motherboard devices
Set oWMISrvEx = GetObject("winmgmts:root/CIMV2")
Set oWMIObjSet = oWMISrvEx.ExecQuery(sWQL)
intRow = 2
strRow = Str(intRow)

ThisWorkbook.Sheets("OnBoardDevices").Range("A1").Value = "Name"
ThisWorkbook.Sheets("OnBoardDevices").Cells(1, 1).Font.Bold = True

ThisWorkbook.Sheets("OnBoardDevices").Range("B1").Value = "Value"
ThisWorkbook.Sheets("OnBoardDevices").Cells(1, 2).Font.Bold = True

For Each oWMIObjEx In oWMIObjSet

For Each oWMIProp In oWMIObjEx.Properties_
If Not IsNull(oWMIProp.Value) Then
If IsArray(oWMIProp.Value) Then
For n = LBound(oWMIProp.Value) To UBound(oWMIProp.Value)
Debug.Print oWMIProp.Name & "(" & n & ")", oWMIProp.Value(n)
ThisWorkbook.Sheets("OnBoardDevices").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).Value = oWMIProp.Value(n)
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
Next
Else
ThisWorkbook.Sheets("OnBoardDevices").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("OnBoardDevices").Range("B" & Trim(strRow)).Value = oWMIProp.Value
ThisWorkbook.Sheets("OnBoardDevices").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
End If
End If
Next
'End If
Next
End Sub
''******************************************************************************************************************************
Sub PrinterWMI()

sWQL = "Select * From Win32_Printer"    'Installed Printers
Set oWMISrvEx = GetObject("winmgmts:root/CIMV2")
Set oWMIObjSet = oWMISrvEx.ExecQuery(sWQL)
intRow = 2
strRow = Str(intRow)

ThisWorkbook.Sheets("Printer").Range("A1").Value = "Name"
ThisWorkbook.Sheets("Printer").Cells(1, 1).Font.Bold = True

ThisWorkbook.Sheets("Printer").Range("B1").Value = "Value"
ThisWorkbook.Sheets("Printer").Cells(1, 2).Font.Bold = True

For Each oWMIObjEx In oWMIObjSet

For Each oWMIProp In oWMIObjEx.Properties_
If Not IsNull(oWMIProp.Value) Then
If IsArray(oWMIProp.Value) Then
For n = LBound(oWMIProp.Value) To UBound(oWMIProp.Value)
Debug.Print oWMIProp.Name & "(" & n & ")", oWMIProp.Value(n)
ThisWorkbook.Sheets("Printer").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).Value = oWMIProp.Value(n)
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
Next
Else
ThisWorkbook.Sheets("Printer").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Printer").Range("B" & Trim(strRow)).Value = oWMIProp.Value
ThisWorkbook.Sheets("Printer").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
End If
End If
Next
'End If
Next
End Sub
''******************************************************************************************************************************
Sub OperatingWMI()

sWQL = "Select * From Win32_OperatingSystem"    'Which Version of Windows with Serial Number
Set oWMISrvEx = GetObject("winmgmts:root/CIMV2")
Set oWMIObjSet = oWMISrvEx.ExecQuery(sWQL)
intRow = 2
strRow = Str(intRow)

ThisWorkbook.Sheets("Operating System").Range("A1").Value = "Name"
ThisWorkbook.Sheets("Operating System").Cells(1, 1).Font.Bold = True

ThisWorkbook.Sheets("Operating System").Range("B1").Value = "Value"
ThisWorkbook.Sheets("Operating System").Cells(1, 2).Font.Bold = True

For Each oWMIObjEx In oWMIObjSet

For Each oWMIProp In oWMIObjEx.Properties_
If Not IsNull(oWMIProp.Value) Then
If IsArray(oWMIProp.Value) Then
For n = LBound(oWMIProp.Value) To UBound(oWMIProp.Value)
Debug.Print oWMIProp.Name & "(" & n & ")", oWMIProp.Value(n)
ThisWorkbook.Sheets("Operating System").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).Value = oWMIProp.Value(n)
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
Next
Else
ThisWorkbook.Sheets("Operating System").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Operating System").Range("B" & Trim(strRow)).Value = oWMIProp.Value
ThisWorkbook.Sheets("Operating System").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
End If
End If
Next
'End If
Next
End Sub
''******************************************************************************************************************************
Sub SoftwareWMI()

sWQL = "Select * From Win32_Product"    'Installed Software
Set oWMISrvEx = GetObject("winmgmts:root/CIMV2")
Set oWMIObjSet = oWMISrvEx.ExecQuery(sWQL)
intRow = 2
strRow = Str(intRow)

ThisWorkbook.Sheets("Software").Range("A1").Value = "Name"
ThisWorkbook.Sheets("Software").Cells(1, 1).Font.Bold = True

ThisWorkbook.Sheets("Software").Range("B1").Value = "Value"
ThisWorkbook.Sheets("Software").Cells(1, 2).Font.Bold = True

For Each oWMIObjEx In oWMIObjSet

For Each oWMIProp In oWMIObjEx.Properties_
If Not IsNull(oWMIProp.Value) Then
If IsArray(oWMIProp.Value) Then
For n = LBound(oWMIProp.Value) To UBound(oWMIProp.Value)
Debug.Print oWMIProp.Name & "(" & n & ")", oWMIProp.Value(n)
ThisWorkbook.Sheets("Software").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).Value = oWMIProp.Value(n)
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
Next
Else
ThisWorkbook.Sheets("Software").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Software").Range("B" & Trim(strRow)).Value = oWMIProp.Value
ThisWorkbook.Sheets("Software").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
End If
End If
Next
'End If
Next
End Sub
''******************************************************************************************************************************
Sub ServicesWMI()

sWQL = "Select * From Win32_BaseService"    'List services running (or stopped) on any PC along with the service's path and file name
Set oWMISrvEx = GetObject("winmgmts:root/CIMV2")
Set oWMIObjSet = oWMISrvEx.ExecQuery(sWQL)
intRow = 2
strRow = Str(intRow)

ThisWorkbook.Sheets("Services").Range("A1").Value = "Name"
ThisWorkbook.Sheets("Services").Cells(1, 1).Font.Bold = True

ThisWorkbook.Sheets("Services").Range("B1").Value = "Value"
ThisWorkbook.Sheets("Services").Cells(1, 2).Font.Bold = True

For Each oWMIObjEx In oWMIObjSet

For Each oWMIProp In oWMIObjEx.Properties_
If Not IsNull(oWMIProp.Value) Then
If IsArray(oWMIProp.Value) Then
For n = LBound(oWMIProp.Value) To UBound(oWMIProp.Value)
Debug.Print oWMIProp.Name & "(" & n & ")", oWMIProp.Value(n)
ThisWorkbook.Sheets("Services").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).Value = oWMIProp.Value(n)
ThisWorkbook.Sheets("Network").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
Next
Else
ThisWorkbook.Sheets("Services").Range("A" & Trim(strRow)).Value = oWMIProp.Name
ThisWorkbook.Sheets("Services").Range("B" & Trim(strRow)).Value = oWMIProp.Value
ThisWorkbook.Sheets("Services").Range("B" & Trim(strRow)).HorizontalAlignment = xlLeft
intRow = intRow + 1
strRow = Str(intRow)
End If
End If
Next
'End If
Next
End Sub

