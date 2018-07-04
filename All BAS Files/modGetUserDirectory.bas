Attribute VB_Name = "modGetUserDirectory"
Option Explicit
Option Compare Text

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modGetUserDirectory
' By Chip Pearson, chip@cpearson.com , www.cpearson.com
'
' This module contains two procedures,
'   GetUserProfileFolder        which returns the folder in which the user's special folders (e.g.,
'                                   "My Documents" or "Recent") are stored.
'
'   GetSpecialFolder            which returns a specific folder for the current user (e.g.,
'                                   "My Documents" or "Recent).
'
' These functions are used to retrieve folder names that are specific the current user.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function CloseHandle Lib "kernel32" _
   (ByVal hObject As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long

Private Declare Function GetCurrentThread Lib "kernel32" () As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function OpenProcessToken Lib "advapi32.dll" ( _
    ByVal ProcessHandle As Long, _
    ByVal DesiredAccess As Long, _
    ByRef TokenHandle As Long) As Long

Private Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" ( _
    ByVal HWnd As Long, _
    ByVal csidl As Long, _
    ByVal hToken As Long, _
    ByVal dwFlags As Long, _
    ByVal pszPath As String) As Long

Private Declare Function GetUserProfileDirectory Lib "userenv.dll" Alias "GetUserProfileDirectoryA" ( _
    ByVal hToken As Long, _
    ByVal lpProfileDir As String, _
    ByRef lpcchSize As Long) As Long

Private Declare Function FormatMessage Lib "kernel32" _
    Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    ByVal lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    ByRef Arguments As Long) As Long




'''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Misc Constants
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const MAX_PATH = 260&
Private Const S_OK = 0&
Private Const E_INVALIDARG As Long = &H80070057
Private Const S_FALSE As Long = &H1 ' odd but true that S_FALSE would be 1.


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Used By OpenProcessToken and OpenProcess
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const TOKEN_QUERY As Long = &H8
Private Const TOKEN_QUERY_SOURCE As Long = &H10
Private Const READ_CONTROL As Long = &H20000
Private Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Private Const TOKEN_READ As Long = (STANDARD_RIGHTS_READ Or TOKEN_QUERY)
Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_VM_READ As Long = (&H10)
Private Const PROCESS_QUERY_INFORMATION As Long = (&H400)



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' used by FormatMessage
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Const FORMAT_MESSAGE_TEXT_LEN = &HA0


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CSIDL Constants of various folder names.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const CSIDL_ADMINTOOLS As Long = &H30
Public Const CSIDL_ALTSTARTUP As Long = &H1D
Public Const CSIDL_APPDATA As Long = &H1A
Public Const CSIDL_BITBUCKET As Long = &HA
Public Const CSIDL_COMMON_ADMINTOOLS As Long = &H2F
Public Const CSIDL_COMMON_ALTSTARTUP As Long = &H1E
Public Const CSIDL_COMMON_APPDATA As Long = &H23
Public Const CSIDL_COMMON_DESKTOPDIRECTORY As Long = &H19
Public Const CSIDL_COMMON_DOCUMENTS As Long = &H2E
Public Const CSIDL_COMMON_FAVORITES As Long = &H1F
Public Const CSIDL_COMMON_PROGRAMS As Long = &H17
Public Const CSIDL_COMMON_STARTMENU As Long = &H16
Public Const CSIDL_COMMON_STARTUP As Long = &H18
Public Const CSIDL_COMMON_TEMPLATES As Long = &H2D
Public Const CSIDL_CONNECTIONS As Long = &H31
Public Const CSIDL_CONTROLS As Long = &H3
Public Const CSIDL_COOKIES As Long = &H21
Public Const CSIDL_DESKTOP As Long = &H0
Public Const CSIDL_DESKTOPDIRECTORY As Long = &H10
Public Const CSIDL_DRIVES As Long = &H11
Public Const CSIDL_FAVORITES As Long = &H6
Public Const CSIDL_FLAG_CREATE As Long = &H8000
Public Const CSIDL_FLAG_DONT_VERIFY As Long = &H4000
Public Const CSIDL_FLAG_MASK As Long = &HFF00&
Public Const CSIDL_FLAG_PFTI_TRACKTARGET As Long = CSIDL_FLAG_DONT_VERIFY
Public Const CSIDL_FONTS As Long = &H14
Public Const CSIDL_HISTORY As Long = &H22
Public Const CSIDL_INTERNET As Long = &H1
Public Const CSIDL_INTERNET_CACHE As Long = &H20
Public Const CSIDL_LOCAL_APPDATA As Long = &H1C
Public Const CSIDL_MYPICTURES As Long = &H27
Public Const CSIDL_NETHOOD As Long = &H13
Public Const CSIDL_NETWORK As Long = &H12
Public Const CSIDL_PERSONAL As Long = &H5   ' My Documents
Public Const CSIDL_MY_DOCUMENTS As Long = &H5
Public Const CSIDL_PRINTERS As Long = &H4
Public Const CSIDL_PRINTHOOD As Long = &H1B
Public Const CSIDL_PROFILE As Long = &H28
Public Const CSIDL_PROGRAM_FILES As Long = &H26
Public Const CSIDL_PROGRAM_FILES_COMMON As Long = &H2B
Public Const CSIDL_PROGRAM_FILES_COMMONX86 As Long = &H2C
Public Const CSIDL_PROGRAM_FILESX86 As Long = &H2A
Public Const CSIDL_PROGRAMS As Long = &H2
Public Const CSIDL_RECENT As Long = &H8
Public Const CSIDL_SENDTO As Long = &H9
Public Const CSIDL_STARTMENU As Long = &HB
Public Const CSIDL_STARTUP As Long = &H7
Public Const CSIDL_SYSTEM As Long = &H25
Public Const CSIDL_SYSTEMX86 As Long = &H29
Public Const CSIDL_TEMPLATES As Long = &H15
Public Const CSIDL_WINDOWS As Long = &H24
Public Const PRIV_PAL_FUNC As Long = &H0
Public Const CSIR_FUNC As Long = (PRIV_PAL_FUNC Or &HD)


Public Function GetSpecialFolder(FolderCSIDL As Long) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' F_7_AB_1_GetSpecialFolder
' This returns the requisted folder name from the user's profile directory. A_7_AB_1_FolderCSIDL must be
' one of the constants beginning with CSIDL listed above. Otherwise, and INVALID ARGUMENT error will
' occur.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim HWnd As Long
Dim Path As String
Dim Res As Long
Dim ErrNumber As Long
Dim ErrText As String

''''''''''''''''''''''''''''''''''''''''''''
' initialize the path variable
''''''''''''''''''''''''''''''''''''''''''''
Path = String$(MAX_PATH, vbNullChar)

''''''''''''''''''''''''''''''''''''''''''''
' get the folder name
''''''''''''''''''''''''''''''''''''''''''''
Res = SHGetFolderPath(HWnd:=0&, _
                        csidl:=FolderCSIDL, _
                        hToken:=0&, _
                        dwFlags:=0&, _
                        pszPath:=Path)
Select Case Res
    Case S_OK
        Path = TrimToNull(Text:=Path)
        GetSpecialFolder = Path
    Case S_FALSE
        MsgBox "The folder code is valid but the folder does not exist."
        GetSpecialFolder = vbNullString
    Case E_INVALIDARG
        MsgBox "The value of FolderCSIDL is not valid."
        GetSpecialFolder = vbNullString
    Case Else
        ErrNumber = Err.LastDllError
        ErrText = GetSystemErrorMessageText(Res)
        MsgBox "An error occurred." & vbCrLf & _
            "System Error: " & CStr(ErrNumber) & vbCrLf & _
            "Description:  " & ErrText
End Select

End Function


Public Function GetUserProfileFolder() As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetUserDirectory
' Return the root directory of the current user's profile. Subfolder of this folder include
'       Application Data
'       Cookies
'       Desktop
'       Favorites
'       Local Settings
'       My Documents
'       NetHood
'       PrintHood
'       Recent
'       SendTo
'       Start Menu
'       Templates
'       UserData
'       Windows
'
' Use GetSpecialFolder above to retrieve the full path name of these folders.
'
' Note: This folder name can also be retrieved with the Environ function:
'
'     Dim UserProfileFolderAs String
'     UserProfileFolder = Environ("UserProfile")
'
' However, I have encountered situations in which an existing Environment variable
' is not accessible by name, which would cause an invalid result in the code above.
' Using the API code will always work.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim Res As Long
Dim CurrentProcessHandle As Long
Dim TokenHandle As Long
Dim UserProfileDirectory As String
Dim LLen As Long
Dim Pos As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Initialize the string to receive the folder name
''''''''''''''''''''''''''''''''''''''''''''''''''''
UserProfileDirectory = String(MAX_PATH, " ")
LLen = Len(UserProfileDirectory)

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get the pseudo-handle of the current process
''''''''''''''''''''''''''''''''''''''''''''''''''''
CurrentProcessHandle = GetCurrentProcess()

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Open the access token of the process
''''''''''''''''''''''''''''''''''''''''''''''''''''
Res = OpenProcessToken(CurrentProcessHandle, TOKEN_READ, TokenHandle)
If Res = 0 Then
    MsgBox "ERROR OpenProcessToken   " & CStr(Err.LastDllError) & " " & GetSystemErrorMessageText(Err.LastDllError)
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get the user's directory
''''''''''''''''''''''''''''''''''''''''''''''''''''
Res = GetUserProfileDirectory(TokenHandle, UserProfileDirectory, LLen)
If Res = 0 Then
    CloseHandle CurrentProcessHandle
    CloseHandle TokenHandle
    MsgBox "ERROR GetUserProfileDirectory   " & CStr(Err.LastDllError) & " " & GetSystemErrorMessageText(Err.LastDllError)
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Trim to null char
''''''''''''''''''''''''''''''''''''''''''''''''''''
UserProfileDirectory = TrimToNull(Text:=UserProfileDirectory)


''''''''''''''''''''''''''''''''''''''''''''''''''''
'Close handles
''''''''''''''''''''''''''''''''''''''''''''''''''''
CloseHandle CurrentProcessHandle
CloseHandle TokenHandle

''''''''''''''''''''''''''''''''''''''''''''''''''''
' Return the result
''''''''''''''''''''''''''''''''''''''''''''''''''''
GetUserProfileFolder = UserProfileDirectory

End Function


Private Function GetSystemErrorMessageText(ErrorNumber As Long) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetSystemErrorMessageText
'
' This function gets the system error message text that corresponds
' to the error code parameter ErrorCode. This value is the value returned
' by Err.LastDLLError or by GetLastError, or occasionally as the returned
' result of a Windows API function.
'
' These are NOT the error numbers returned by Err.Number (for these
' errors, use Err.Description to get the description of the error).
'
' In general, you should use Err.LastDllError rather than GetLastError
' because under some circumstances the value of GetLastError will be
' reset to 0 before the value is returned to VBA. Err.LastDllError will
' always reliably return the last error number raised in an API function.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim ErrorText As String
Dim TextLen As Long
Dim FormatMessageResult As Long
Dim LangID As Long

''''''''''''''''''''''''''''''''
' initialize the variables
''''''''''''''''''''''''''''''''
LangID = 0& 'default language
ErrorText = String$(FORMAT_MESSAGE_TEXT_LEN, vbNullChar)
TextLen = FORMAT_MESSAGE_TEXT_LEN

' Call FormatMessage to get the text of the error message text
' associated with ErrorNumber.
FormatMessageResult = FormatMessage( _
                        dwFlags:=FORMAT_MESSAGE_FROM_SYSTEM Or _
                                 FORMAT_MESSAGE_IGNORE_INSERTS, _
                        lpSource:=0&, _
                        dwMessageId:=ErrorNumber, _
                        dwLanguageId:=LangID, _
                        lpBuffer:=ErrorText, _
                        nSize:=TextLen, _
                        Arguments:=0&)
If FormatMessageResult = 0& Then
    ' An error occured. Display the error number, but
    ' don't call GetSystemErrorMessageText to get the
    ' text, which would likely cause the error again,
    ' getting us into a loop.

    MsgBox "An error occurred with the FormatMessage" & _
           " API functiopn call. Error: " & _
           CStr(Err.LastDllError) & _
           " Hex(" & Hex(Err.LastDllError) & ")."
    GetSystemErrorMessageText = vbNullString
    Exit Function
End If

' If FormatMessageResult is not zero, it is the number
' of characters placed in the ErrorText variable.
' Take the left FormatMessageResult characters and
' return that text.
ErrorText = Left$(ErrorText, FormatMessageResult)
GetSystemErrorMessageText = ErrorText

End Function

Private Function TrimToNull(Text As String) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' TrimToNull
' This function returns the portion of Text that is to the left of the vbNullChar
' character (same as Chr(0)). Typically, this function is used with strings
' populated by Windows API procedures. It is generally not used for
' native VB Strings.
' If vbNullChar is not found, the entire Text string is returned.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Pos As Integer
    Pos = InStr(1, Text, vbNullChar)
    If Pos > 0 Then
        TrimToNull = Left(Text, Pos - 1)
    Else
        TrimToNull = Text
    End If

End Function




