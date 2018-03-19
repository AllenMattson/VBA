Attribute VB_Name = "GetPrinters"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modListPrinters
' By Chip Pearson, chip@cpearson.com  www.cpearson.com
' Created 22-Sept-2012
' This provides a function named GetPrinterFullNames that
' returns a String array, each element of which is the name
' of a printer installed on the machine.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKCU = HKEY_CURRENT_USER
Private Const KEY_QUERY_VALUE = &H1&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_MORE_DATA = 234

Private Declare Function RegCloseKey Lib "advapi32.dll" ( _
    ByVal HKey As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32" _
    Alias "RegOpenKeyExA" ( _
    ByVal HKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long

Private Declare Function RegEnumValue Lib "advapi32.dll" _
    Alias "RegEnumValueA" ( _
    ByVal HKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpValueName As String, _
    lpcbValueName As Long, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Byte, _
    lpcbData As Long) As Long

    Sub GetPrinters()
        Dim Printers() As String
        Dim N As Long
        Dim S As String
        Dim i As Integer
        Printers = GetPrinterFullNames()
        For N = LBound(Printers) To UBound(Printers)
            i = 1
            Cells((N + 1), 1).Value = N
            Cells(N + 1, 2).Value = Printers(N)
            S = S & Printers(N) & vbNewLine
        Next N
        MsgBox S, vbOKOnly, "Printers"
    End Sub

Public Function GetPrinterFullNames() As String()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetPrinterFullNames
' By Chip Pearson, chip@cpearson.com, www.cpearson.com
' Returns an array of printer names, where each printer name
' is the device name followed by the port name. The value can
' be used to assign a printer to the ActivePrinter property.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Printers() As String ' array of names to be returned
Dim PNdx As Long    ' index into Printers()
Dim HKey As Long    ' registry key handle
Dim Res As Long     ' result of API calls
Dim Ndx As Long     ' index for RegEnumValue
Dim ValueName As String ' name of each value in the printer key
Dim ValueNameLen As Long    ' length of ValueName
Dim DataType As Long        ' registry value data type
Dim ValueValue() As Byte    ' byte array of registry value value
Dim ValueValueS As String   ' ValueValue converted to String
Dim CommaPos As Long        ' position of comma character in ValueValue
Dim ColonPos As Long        ' position of colon character in ValueValue
Dim M As Long               ' string index

' registry key in HCKU listing printers
Const PRINTER_KEY = "Software\Microsoft\Windows NT\CurrentVersion\Devices"

PNdx = 0
Ndx = 0
' assume printer name is less than 256 characters
ValueName = String$(256, Chr(0))
ValueNameLen = 255
' assume the port name is less than 1000 characters
ReDim ValueValue(0 To 999)
' assume there are less than 1000 printers installed
ReDim Printers(1 To 1000)

' open the key whose values enumerate installed printers
Res = RegOpenKeyEx(HKCU, PRINTER_KEY, 0&, _
    KEY_QUERY_VALUE, HKey)
' start enumeration loop of printers
Res = RegEnumValue(HKey, Ndx, ValueName, _
    ValueNameLen, 0&, DataType, ValueValue(0), 1000)
' loop until all values have been enumerated
Do Until Res = ERROR_NO_MORE_ITEMS
    M = InStr(1, ValueName, Chr(0))
    If M > 1 Then
        ' clean up the ValueName
        ValueName = Left(ValueName, M - 1)
    End If
    ' find position of a comma and colon in the port name
    CommaPos = InStr(1, ValueValue, ",")
    ColonPos = InStr(1, ValueValue, ":")
    ' ValueValue byte array to ValueValueS string
    On Error Resume Next
    ValueValueS = Mid(ValueValue, CommaPos + 1, ColonPos - CommaPos)
    On Error GoTo 0
    ' next slot in Printers
    PNdx = PNdx + 1
    Printers(PNdx) = ValueName & " on " & ValueValueS
    ' reset some variables
    ValueName = String(255, Chr(0))
    ValueNameLen = 255
    ReDim ValueValue(0 To 999)
    ValueValueS = vbNullString
    ' tell RegEnumValue to get the next registry value
    Ndx = Ndx + 1
    ' get the next printer
    Res = RegEnumValue(HKey, Ndx, ValueName, ValueNameLen, _
        0&, DataType, ValueValue(0), 1000)
    ' test for error
    If (Res <> 0) And (Res <> ERROR_MORE_DATA) Then
        Exit Do
    End If
Loop
' shrink Printers down to used size
ReDim Preserve Printers(1 To PNdx)
Res = RegCloseKey(HKey)
GetPrinterFullNames = Printers
End Function

