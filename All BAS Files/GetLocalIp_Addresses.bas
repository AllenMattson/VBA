Attribute VB_Name = "GetLocalIp_Addresses"
Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long

' Returns an array with the local IP addresses (as strings).
' Author: Christian d'Heureuse, www.source-code.biz
Public Function GetIpAddrTable()
   Dim Buf(0 To 511) As Byte
   Dim BufSize As Long: BufSize = UBound(Buf) + 1
   Dim rc As Long
   rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
   If rc <> 0 Then Err.Raise vbObjectError, , "GetIpAddrTable failed with return value " & rc
   Dim NrOfEntries As Integer: NrOfEntries = Buf(1) * 256 + Buf(0)
   If NrOfEntries = 0 Then GetIpAddrTable = Array(): Exit Function
   ReDim IpAddrs(0 To NrOfEntries - 1) As String
   Dim i As Integer
   For i = 0 To NrOfEntries - 1
      Dim j As Integer, s As String: s = ""
      For j = 0 To 3: s = s & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j): Next
      IpAddrs(i) = s
      Next
   GetIpAddrTable = IpAddrs
   End Function



' Test program for GetIpAddrTable.
Public Sub Test()
   Dim IpAddrs
   IpAddrs = GetIpAddrTable
   Debug.Print "Nr of IP addresses: " & UBound(IpAddrs) - LBound(IpAddrs) + 1
   Dim i As Integer
   For i = LBound(IpAddrs) To UBound(IpAddrs)
      Debug.Print IpAddrs(i)
      Next
   End Sub
