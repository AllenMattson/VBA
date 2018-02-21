Attribute VB_Name = "Module1"
Option Explicit

#If VBA7 And Win64 Then
    
    Private Declare PtrSafe Function RegOpenKeyA Lib "ADVAPI32.DLL" _
        (ByVal hKey As LongPtr, ByVal lpSubKey As String, _
        phkResult As LongPtr) As Long
        
    Private Declare PtrSafe Function RegCloseKey Lib "ADVAPI32.DLL" _
        (ByVal hKey As LongPtr) As Long
    
    Private Declare PtrSafe Function RegSetValueExA Lib "ADVAPI32.DLL" _
        (ByVal hKey As LongPtr, ByVal sValueName As String, _
        ByVal dwReserved As Long, ByVal dwType As Long, _
        ByVal sValue As String, ByVal dwSize As Long) As Long
    
    Private Declare PtrSafe Function RegCreateKeyA Lib "ADVAPI32.DLL" _
        (ByVal hKey As LongPtr, ByVal sSubKey As String, _
        ByRef hkeyResult As LongPtr) As Long
        
    Private Declare PtrSafe Function RegQueryValueExA Lib "ADVAPI32.DLL" _
        (ByVal hKey As LongPtr, ByVal sValueName As String, _
        ByVal dwReserved As Long, ByRef lValueType As Long, _
        ByVal sValue As String, ByRef lResultLen As Long) As Long
#Else
    Private Declare Function RegOpenKeyA Lib "ADVAPI32.DLL" _
        (ByVal hKey As Long, ByVal sSubKey As String, _
        ByRef hkeyResult As Long) As Long
    
    Private Declare Function RegCloseKey Lib "ADVAPI32.DLL" _
        (ByVal hKey As Long) As Long
    
    Private Declare Function RegSetValueExA Lib "ADVAPI32.DLL" _
        (ByVal hKey As Long, ByVal sValueName As String, _
        ByVal dwReserved As Long, ByVal dwType As Long, _
        ByVal sValue As String, ByVal dwSize As Long) As Long
    
    Private Declare Function RegCreateKeyA Lib "ADVAPI32.DLL" _
        (ByVal hKey As Long, ByVal sSubKey As String, _
        ByRef hkeyResult As Long) As Long
    
    Private Declare Function RegQueryValueExA Lib "ADVAPI32.DLL" _
        (ByVal hKey As Long, ByVal sValueName As String, _
        ByVal dwReserved As Long, ByRef lValueType As Long, _
        ByVal sValue As String, ByRef lResultLen As Long) As Long
#End If

Sub UpdateRegistryWithTime()
Attribute UpdateRegistryWithTime.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim RootKey As String
    Dim Path As String
    Dim RegEntry As String
    Dim RegVal As Date
    Dim LastTime As String
    Dim Msg As String
    
    RootKey = "hkey_current_user"
    Path = "software\microsoft\office\14.0\excel\LastStarted"
    RegEntry = "DateTime"
    RegVal = Now()
    
    LastTime = GetRegistry(RootKey, Path, RegEntry)
    Select Case LastTime
        Case "Not Found"
            Msg = "This routine has not been executed before."
        Case Else
            Msg = "This routine was lasted executed: " & LastTime
    End Select
    Msg = Msg & Chr(13) & Chr(13)
    
    Select Case WriteRegistry(RootKey, Path, RegEntry, RegVal)
        Case True
            Msg = Msg & "The registry has been updated with the current date and time."
        Case False
            Msg = Msg & "An error occured writing to the registry..."
    End Select
    MsgBox Msg, vbInformation, "Registry Demo"
End Sub

Private Function GetRegistry(Key, Path, ByVal ValueName As String)
Attribute GetRegistry.VB_ProcData.VB_Invoke_Func = " \n14"
'  Reads a value from the Windows Registry

#If VBA7 And Win64 Then
    Dim TheKey As LongPtr
    Dim hKey As LongPtr
#Else
    Dim TheKey As Long
    Dim hKey As Long
#End If
    
    Dim lValueType As Long
    Dim sResult As String
    Dim lResultLen As Long
    Dim ResultLen As Long
    Dim x


    TheKey = -99
    Select Case UCase(Key)
        Case "HKEY_CLASSES_ROOT": TheKey = &H80000000
        Case "HKEY_CURRENT_USER": TheKey = &H80000001
        Case "HKEY_LOCAL_MACHINE": TheKey = &H80000002
        Case "HKEY_USERS": TheKey = &H80000003
        Case "HKEY_CURRENT_CONFIG": TheKey = &H80000004
        Case "HKEY_DYN_DATA": TheKey = &H80000005
    End Select
    
'   Exit if key is not found
    If TheKey = -99 Then
        GetRegistry = "Not Found"
        Exit Function
    End If

    If RegOpenKeyA(TheKey, Path, hKey) <> 0 Then _
        x = RegCreateKeyA(TheKey, Path, hKey)
    
    sResult = Space(100)
    lResultLen = 100
    
    x = RegQueryValueExA(hKey, ValueName, 0, lValueType, _
    sResult, lResultLen)
        
    Select Case x
        Case 0: GetRegistry = Left(sResult, lResultLen - 1)
        Case Else: GetRegistry = "Not Found"
    End Select
    
    RegCloseKey hKey
End Function

Private Function WriteRegistry(ByVal Key As String, _
    ByVal Path As String, ByVal entry As String, _
    ByVal value As String)
Attribute WriteRegistry.VB_ProcData.VB_Invoke_Func = " \n14"
    
#If VBA7 And Win64 Then
    Dim TheKey As LongPtr
    Dim hKey As LongPtr
#Else
    Dim TheKey As Long
    Dim hKey As Long
#End If
    
    
    Dim lValueType As Long
    Dim sResult As String
    Dim lResultLen As Long
    Dim x
    
   
    TheKey = -99
    Select Case UCase(Key)
        Case "HKEY_CLASSES_ROOT": TheKey = &H80000000
        Case "HKEY_CURRENT_USER": TheKey = &H80000001
        Case "HKEY_LOCAL_MACHINE": TheKey = &H80000002
        Case "HKEY_USERS": TheKey = &H80000003
        Case "HKEY_CURRENT_CONFIG": TheKey = &H80000004
        Case "HKEY_DYN_DATA": TheKey = &H80000005
    End Select
    
'   Exit if key is not found
    If TheKey = -99 Then
        WriteRegistry = False
        Exit Function
    End If

'   Make sure  key exists
    If RegOpenKeyA(TheKey, Path, hKey) <> 0 Then
        x = RegCreateKeyA(TheKey, Path, hKey)
    End If

    x = RegSetValueExA(hKey, entry, 0, 1, value, Len(value) + 1)
    If x = 0 Then WriteRegistry = True Else WriteRegistry = False
End Function


Sub Wallpaper()
    Dim RootKey As String
    Dim Path As String
    Dim RegEntry As String
    RootKey = "hkey_current_user"
    Path = "Control Panel\Desktop"
    RegEntry = "Wallpaper"
    MsgBox GetRegistry(RootKey, Path, RegEntry), vbInformation, Path & "\RegEntry"
End Sub
