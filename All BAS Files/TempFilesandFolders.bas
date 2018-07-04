Attribute VB_Name = "TempFilesandFolders"
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modTempFilesAndFolders
' By Chip Pearson, chip@cpearson.com, www.cpearson.com
'
' This module contains the procedures GetTempFile, GetTempFolderName, and GetTemporaryFolderName.
' is in each procedure.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''
' Maximum Length Of Full File Name
'''''''''''''''''''''''''''''''''''
Private Const MAX_PATH = 260 ' Windows Standard, from VC++ StdLib.h

'''''''''''''''''''''''''''''''''''
' used by FormatMessage
'''''''''''''''''''''''''''''''''''
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_TEXT_LEN = &HA0 ' from ERRORS.H C++ include file.

'''''''''''''''''''''''''''''''''''
' Win API Declares
'''''''''''''''''''''''''''''''''''
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    ByRef Arguments As Long) As Long

Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" ( _
    ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long
    
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
    ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long

Private Declare Function PathGetCharType Lib "shlwapi.dll" _
    Alias "PathGetCharTypeA" ( _
    ByVal ch As Byte) As Long

Public Function GetTemporaryFolderName(Optional Create As Boolean = False) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetTemporaryFolderName
' This function returns the name of a temporary folder name. The folder will be
' in the user's designated temp folder. If Create is True, the folder will
' be created.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim FName As String
Dim FileName As String
Dim TempFolderName As String
Dim Pos As Integer

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get a temp file name with no extension, located in the
' user's system-specified temporary folder.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
FName = GetTempFile(vbNullString, vbNullString, " ", False)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Find the location of the last "\" character.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Pos = InStrRev(FName, "\", -1, vbTextCompare)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get the filename (without the path) of the temp file.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
FileName = Mid(FName, Pos + 1)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get the user's system-specified temp folder name
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
TempFolderName = GetTempFolderName(IncludeTrailingSlash:=True)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' append FolderName to the full folder fname
TempFolderName = TempFolderName & FileName


''''''''''''''''''''''''''''''''''''''''
' Create the folder is requested.
''''''''''''''''''''''''''''''''''''''''
If Create = True Then
    On Error Resume Next
    Err.Clear
    MkDir TempFolderName
    If Err.Number <> 0 Then
        MsgBox "An error occurred creating folder '" & TempFolderName & "'" & vbCrLf & _
            "Err: " & CStr(Err.Number) & vbCrLf & _
            "Description: " & Err.Description
        GetTemporaryFolderName = vbNullString
        Exit Function
    End If
End If

    
''''''''''''''''''''''''''''''''''''''''
' return the result
''''''''''''''''''''''''''''''''''''''''
GetTemporaryFolderName = TempFolderName

End Function



Public Function GetTempFolderName( _
    Optional IncludeTrailingSlash As Boolean = False) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetTempFolder
' This procedure returns a temporary folder in which the application may safely
' store temporary files. Returns the name of the folder or vbNullString if an error occurred.
' The argument IncludeTrailingSlash indicates whether to include a trailing slash
' at the end of the folder name.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim TempPath As String
Dim Length As Long
Dim Result As Long
Dim ErrorNumber As Long
Dim ErrorText As String

''''''''''''''''''''''''''''''''''''''
' Initialize the variables
''''''''''''''''''''''''''''''''''''''
TempPath = String$(MAX_PATH, " ")
Length = MAX_PATH

'''''''''''''''''''''''''''''''''''''''''''''''''
' Get the Temporary Path using GetTempPath.
'''''''''''''''''''''''''''''''''''''''''''''''''
Result = GetTempPath(Length, TempPath)
If Result = 0 Then
    '''''''''''''''''''''''''''''''''''''
    ' An error occurred
    '''''''''''''''''''''''''''''''''''''
    ErrorNumber = Err.LastDllError
    ErrorText = GetSystemErrorMessageText(ErrorNumber)
    MsgBox "An error occurred getting the temporary folder" & _
           " from the GetTempFolderName function: " & vbCrLf & _
           "Error: " & CStr(ErrorNumber) & "  " & ErrorText
    GetTempFolderName = vbNullString
    Exit Function
Else
    '''''''''''''''''''''''''''''''''''''''
    ' No error, but the buffer may have
    ' been too small.
    '''''''''''''''''''''''''''''''''''''''
    If Result > Length Then
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' The buffer TempPath was too small to hold the folder name.
        ' This should never happen if MAX_PATH is set to the proper
        ' value.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        MsgBox "The TempPath buffer is too small. It is allocated at " & _
            CStr(Length) & " characters." & vbCrLf & _
            "The required buffer size is: " & CStr(Result) & " characteres.", _
            vbOKOnly, "GetTempFolderName"
        GetTempFolderName = vbNullString
        Exit Function
    End If

    ' trim up the TempPath. It includes a trailing "\"
    TempPath = TrimToNull(Text:=TempPath)
    
    If IncludeTrailingSlash = False Then
        '''''''''''''''''''''''''''''''''''''''''''''''''
        ' If IncludeTrailingSlash is False, get rid of
        ' the trailing slash.
        '''''''''''''''''''''''''''''''''''''''''''''''''
        TempPath = Left(TempPath, Len(TempPath) - 1)
    End If
End If

'''''''''''''''''''''''''''''''''
' return the result
'''''''''''''''''''''''''''''''''
GetTempFolderName = TempPath

End Function


Public Function GetTempFile(Optional InFolder As String = vbNullString, _
                            Optional FileNamePrefix As String = vbNullString, _
                            Optional Extension As String = vbNullString, _
                            Optional CreateFile As Boolean = True)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetTempFileName
'
' This function will return the name of a temporary file, optionally suffixed with the
' string in the Extension variable. It will optionally create the file. (Actually,
' GetTempFileName will create the file, so we will optionally Kill it.)
'
' If InFolder specifies an existing folder, the file will be created in that folder.
' If InFolder specifies a non-existant folder, the procedure will attempt to create
' the folder.
' If InFolder is vbNullString, the procedure will call GetTempFolderName to get
' the folder designated for temporary files.
' InFolder must be a fully qualified path. That is, a folder name begining with a
' network prefix "\\" or containing ":".

' If FileNamePrefix is specified, the file name will begin with the first three
' characters of this string. In this case, FileNamePrefix must be three characters
' with no spaces or illegal file name characters. These are validated with
' PathGetCharType. If FileNamePrefix is vbNullString, the value of C_DEFAULT_PREFIX
' will be used.
' If FileNamePrefix contains spaces or invalid characters, an error occurs.
'
' If Extension is specified, the filename will have that Extension. If must be three
' valid characters (no spaces). The characters are validated with PathGetCharType.
' If Extension is vbNullString the default extension from GetTempFileName ("tmp") is
' used. Do NOT put the period in front of the extension (e.g., use "xls" not ".xls").
' If Extension is a single space, the file name will have no extension.
'
' If CreateFile is omitted or True, the file will be created. If CreateFile is false,
' the file is not created. (Actually, it will be created by GetTempFileName  and then
' KILLed.)
'
' A temp file may be created on in network location.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim PathBuffer As String
Dim Prefix As String
Dim FolderPath As String
Dim Res As Long
Dim FileName As String
Dim ErrorNumber As Long
Dim ErrorText As String
Dim FileNumber As Integer


Const C_DEFAULT_PREFIX = "TMP"
FileName = String$(MAX_PATH, vbNullChar)


If InFolder = vbNullString Then
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' InFolder was an empty string. Call GetTempFolderName
    ' to get a temporary folder name.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    PathBuffer = GetTempFolderName(IncludeTrailingSlash:=True)
Else
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' test to see if we have an absolute path
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If (Left(InFolder, 2) = "\\") Or _
        (InStr(1, InFolder, ":", vbTextCompare) > 0) Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' We have an absolute path. Test whether the folder exists.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Dir(InFolder, vbHidden + vbSystem + vbHidden + _
                         vbNormal + vbDirectory) = vbNullString Then
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            ' InFolder does not exist. Try to create it.
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            On Error Resume Next
            Err.Clear
            MkDir InFolder
            If Err.Number <> 0 Then
                MsgBox "An error occurred creating the '" & InFolder _
                    & "' folder." & vbCrLf & _
                    "Error: " & CStr(Err.Number) & vbCrLf & _
                    "Description: " & Err.Description, vbOKOnly, "GetTempFileName"
                GetTempFile = vbNullString
                Exit Function
            Else
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' MkDir succussfully created the folder. Set PathBuffer to the new
                ' folder name.
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                PathBuffer = InFolder
            End If
        Else
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            ' InFolder exists. Set the PathBuffer variable to InFolder
            '''''''''''''''''''''''''''''''''''''''''''''''''''
            PathBuffer = InFolder
        End If
    Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' We don't have a fully qualified path. Get out with an error message.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        MsgBox "The InFolder parameter to GetTempFile is not an absolute file name.", _
                vbOKOnly, "GetTempFileName"
        GetTempFile = vbNullString
        Exit Function
    End If ' LEFT
End If ' InFolder = vbNullString

''''''''''''''''''''''''''''''''''''''''''
' Ensure we have a '\' at the end of the
' path.
'''''''''''''''''''''''''''''''''''''''''
If Right(PathBuffer, 1) <> "\" Then
    PathBuffer = PathBuffer & "\"
End If

If FileNamePrefix = vbNullString Then
    '''''''''''''''''''''''''''''''''''''''''
    ' FileNamePrefix is empty, use 'tmp'
    '''''''''''''''''''''''''''''''''''''''''
    Prefix = C_DEFAULT_PREFIX
Else
    If IsValidFileNamePrefixOrExtension(Spec:=FileNamePrefix) = False Then
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' FileNamePrefix is invalid. Get out with an error.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        MsgBox "The file name prefix '" & FileNamePrefix & "' is invalid.", _
                            vbOKOnly, "GetTempFileName"
        GetTempFile = vbNullString
        Exit Function
    Else
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' FileNamePrefix is valid.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Prefix = FileNamePrefix
    End If
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Get the temp file name. GetTempFileName will automatically
' create the file. If CreateFile is False, we'll have
' to Kill the file. We set wUnique to 0 to ensure that
' the filename will be unique. This has the side effect
' of creating the file.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Res = GetTempFileName(lpszPath:=PathBuffer, _
                        lpPrefixString:=Prefix, _
                        wUnique:=0, _
                        lpTempFileName:=FileName)
                        
If Res = 0 Then
    ''''''''''''''''''''''''''''
    ' An error occurred. Get out
    ' with an error message.
    ''''''''''''''''''''''''''''
    ErrorNumber = Err.LastDllError
    ErrorText = GetSystemErrorMessageText(ErrorNumber)
    MsgBox "An error occurred with GetTempFileName" & vbCrLf & _
        "Error: " & CStr(ErrorNumber) & vbCrLf & _
        "Description: " & ErrorText, vbOKOnly, "GetTempFileName"
    GetTempFile = vbNullString
    Exit Function
End If

''''''''''''''''''''''''''''''''''''''''''
' GetTempFileName put the file name in the
' FileName variable, ending with a vbNullChar.
' Trim to the the vbNullChar.
'''''''''''''''''''''''''''''''''''''''''''
FileName = TrimToNull(Text:=FileName)



'''''''''''''''''''''''''''''''''''''''''''
' GetTempFileName created a file with an
' extension of "tmp". If Extension was
' specified and is not a null string,
' change the extension to the specified
' extension. We'll use the same validation
' routine as we did for the prefix.
'''''''''''''''''''''''''''''''''''''''''''
If Extension = vbNullString Then
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' If  Extension is vbNullString, use the extension
    ' created by GetTEmpFileName ("tmp"). Test whether
    ' CreateFile is False. If False, we have to kill the
    ' newly created file.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If CreateFile = False Then
        On Error Resume Next
        Kill FileName
    Else
        ''''''''''''''''''''''''''''''''''
        ' CreateFile was true. Leave
        ' the newly created file in place
        ''''''''''''''''''''''''''''''''''
    End If
Else ' Extension is not vbNullString
    If Extension = " " Then
        ''''''''''''''''''''''''''''''''''''
        ' An Extension value of " " indicates
        ' that the filename should have no
        ' extension. First Kill FileName, modify
        ' the variable to have no extension, and then
        ' see if we need to create the file. If CreateFile
        ' if False, don't create the file. If True,
        ' create the file by openning it and then
        ' immmediately close it.
        ''''''''''''''''''''''''''''''''''''
        On Error Resume Next
        Kill FileName
        On Error GoTo 0
        FileName = Left(FileName, Len(FileName) - 4)
        If CreateFile = True Then
            ''''''''''''''''''''''''''''''''''''''''
            ' Create the file by opening it for
            ' output, then immediately closing it.
            ''''''''''''''''''''''''''''''''''''''''
            FileNumber = FreeFile
            Open FileName For Output Access Write As #FileNumber
            Close #FileNumber
        Else
            '''''''''''''''''''''''''''''''''''''''''
            ' CreateFile was false. Since we've already
            ' Killed the file created by GetTempFileName,
            ' do nothing.
            ''''''''''''''''''''''''''''''''''''''''''
        End If
        
            
    Else ' extension is not " "
        
        If IsValidFileNamePrefixOrExtension(Spec:=Extension) = True Then
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' If we have a valid extension, kill the existing filename
            ' and the recreate the file with the new extension.
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            On Error Resume Next
            Kill FileName
            On Error GoTo 0
            FileName = Left(FileName, Len(FileName) - 4) & "." & Extension
            If CreateFile = True Then
                FileNumber = FreeFile
                Open FileName For Output Access Write As #FileNumber
                Close #FileNumber
            Else
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' CreateFile was false. Since we've already killed the
                ' filename created by GetTempFileName, do nothing.
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            End If
        Else
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' The extension was not valid. Display an error and get out.
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            MsgBox "The extension '" & Extension & "' is  not valid.", _
                vbOKOnly, "GetTempFileName"
            GetTempFile = vbNullString
            Exit Function
        End If
        
    End If
    
End If

''''''''''''''''''''''''''''''''''''''''''''
' We were successful. Return the filename.
''''''''''''''''''''''''''''''''''''''''''''
GetTempFile = FileName


End Function

Private Function IsValidFileNamePrefixOrExtension(Spec As String) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsValidFileNamePrefix
' This returns TRUE if Prefix is a valid 3 character filename
' prefix used with GetTempFileName
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Const GCT_INVALID As Long = &H0
Const GCT_SEPARATOR As Long = &H8
Const GCT_WILD As Long = &H4
Const GCT_LFNCHAR As Long = &H1
Const GCT_SHORTCHAR As Long = &H2

Dim Ndx As Long
Dim B As Byte
'''''''''''''''''''''''''''''''''
' Spec contains a space. error.
'''''''''''''''''''''''''''''''''
If InStr(1, Spec, " ") > 0 Then
    IsValidFileNamePrefixOrExtension = False
    Exit Function
End If


'''''''''''''''''''''''''''''''''
' Spec is not 3 chars. error.
'''''''''''''''''''''''''''''''''
If Len(Spec) <> 3 Then
    IsValidFileNamePrefixOrExtension = False
    Exit Function
End If

'''''''''''''''''''''''''''''''''
' Loop through the 3 characters
' of Spec. If we find an
' invalid character, get out with
' a result of False.
'''''''''''''''''''''''''''''''''
For Ndx = 1 To 3
    ''''''''''''''''''''''''''''''''''''''''''
    ' convert the character to a Byte variable
    '''''''''''''''''''''''''''''''''''''''''''
    B = CByte(Asc(Mid(Spec, Ndx, 1)))
        
    ''''''''''''''''''''''''''''''''''''''''''''
    ' test the character B to see if it
    ' is valid in a file name.
    ''''''''''''''''''''''''''''''''''''''''''''
    Select Case PathGetCharType(B)
        Case GCT_INVALID, GCT_SEPARATOR, GCT_WILD
            ''''''''''''''''''''''''''''''''''''''''''
            ' Invalid character, path or drive separator,
            ' or wildcard character. None are allowed.
            ''''''''''''''''''''''''''''''''''''''''''
            IsValidFileNamePrefixOrExtension = False
            Exit Function
        
        Case GCT_LFNCHAR, GCT_SHORTCHAR, GCT_LFNCHAR + GCT_SHORTCHAR
            ''''''''''''''''''''''''''''''''''''''''''''
            ' Character is valid for short file names,
            ' long file names, or both. Do nothing.
            ''''''''''''''''''''''''''''''''''''''''''''
            
        Case Else
            ''''''''''''''''''''''''''''''''''''''''''''
            ' we should never get here, but if we do,
            ' be safe and return False.
            ''''''''''''''''''''''''''''''''''''''''''''
            IsValidFileNamePrefixOrExtension = False
            Exit Function
    End Select
Next Ndx

'''''''''''''''''''''''''''''''''
' If we made it out of the loop,
' the Prefix was valid. Return
' True.
'''''''''''''''''''''''''''''''''
IsValidFileNamePrefixOrExtension = True
    
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

Public Function GetSystemErrorMessageText(ErrorNumber As Long) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetSystemErrorMessageText
'
' This function gets the system error message text that corresponds
' to the error code returned by the GetLastError API function or the
' Err.LastDllError property.
'
' It may be used ONLY for these error codes. These are NOT the error numbers
' returned by Err.Number (for these errors, use Err.Description to get the
' description of the message). The error number MUST be the value returned by
' GetLastError or Err.LastDLLError.
'
' In general, you should use Err.LastDllError rather than GetLastError
' because under some circumstances the value of GetLastError will be
' reset to 0 before the value is returned to VB. Err.LastDllError will
' always reliably return the last error number raised in a DLL.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim ErrorText As String
Dim TextLen As Long
Dim FormatMessageResult As Long
Dim LangID As Long

' initialize the variables
LangID = 0&  'default language
ErrorText = String$(FORMAT_MESSAGE_TEXT_LEN, vbNullChar)
TextLen = Len(ErrorText)
On Error Resume Next
' Call FormatMessage to get the text of the error message
' associated with ErrorNumber.
FormatMessageResult = FormatMessage( _
                dwFlags:=FORMAT_MESSAGE_FROM_SYSTEM Or _
                         FORMAT_MESSAGE_IGNORE_INSERTS, _
                lpSource:=0&, _
                dwMessageId:=ErrorNumber, _
                dwLanguageId:=0&, _
                lpBuffer:=ErrorText, _
                nSize:=TextLen, _
                Arguments:=0&)
On Error GoTo 0
If FormatMessageResult = 0& Then
    ' An error occured. Display the error number, but
    ' don't call GetSystemErrorMessageText  to get the
    ' text, which would likely cause the error again,
    ' getting us into a loop.
    MsgBox "An error occurred with the FormatMessage" & _
        " API functiopn call. Error: " & _
        CStr(Err.LastDllError) & _
        " Hex(" & Hex(Err.LastDllError) & ")."
    GetSystemErrorMessageText = vbNullString
    Exit Function
End If
If FormatMessageResult > 0 Then
    ' success
    ' FormatMessage returned some text. Take the left
    ' FormatMessageResult characters and return that text.
    ErrorText = Left$(ErrorText, FormatMessageResult)
    GetSystemErrorMessageText = ErrorText
Else
    ' an error occurred
    ' Format message didn't return any text.
    ' There is no text description for the specified error.
    GetSystemErrorMessageText = "NO ERROR DESCRIPTION AVAILABLE"
End If


End Function







