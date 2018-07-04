Attribute VB_Name = "modGetSystemErrorMessageText"
Option Explicit
Option Compare Text
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetSystemErrorMessageText
' -------------------------
' By Chp Pearson, www.cpearson.com, chip@cpearson.com
' See www.cpearson.com/Excel/FormatMessage.aspx for
' additional information.
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''
' Used by FormatMessage
'''''''''''''''''''''''''''''''''''
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY  As Long = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE  As Long = &H800
Private Const FORMAT_MESSAGE_FROM_STRING  As Long = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM  As Long = &H1000
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK  As Long = &HFF
Private Const FORMAT_MESSAGE_IGNORE_INSERTS  As Long = &H200
Private Const FORMAT_MESSAGE_TEXT_LEN  As Long = &HA0 ' from VC++ ERRORS.H file

'''''''''''''''''''''''''''''''''''
' Windows API Declare
'''''''''''''''''''''''''''''''''''
Private Declare Function FormatMessage Lib "kernel32" _
    Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    ByVal lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    ByRef Arguments As Long) As Long



Public Function GetSystemErrorMessageText(ErrorNumber As Long) As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetSystemErrorMessageText
'
' This function gets the system error message text that corresponds
' to the error code parameter ErrorNumber. This value is the value returned
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
'
' The function returns vbNullString is an error occurred or if there is
' no error text for the specified error number.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim ErrorText As String
Dim TextLen As Long
Dim FormatMessageResult As Long
Dim LangID As Long

''''''''''''''''''''''''''''''''
' Initialize the variables
''''''''''''''''''''''''''''''''
LangID = 0&   ' Default language
ErrorText = String$(FORMAT_MESSAGE_TEXT_LEN, vbNullChar)
TextLen = FORMAT_MESSAGE_TEXT_LEN

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Call FormatMessage to get the text of the error message text
' associated with ErrorNumber.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    ' An error occured. Display the error number, but
    ' don't call GetSystemErrorMessageText to get the
    ' text, which would likely cause the error again,
    ' getting us into a loop.
    ''''''''''''''''''''''''''''''''''''''''''''''''''
    MsgBox "An error occurred with the FormatMessage" & _
           " API function call." & vbCrLf & _
           "Error: " & CStr(Err.LastDllError) & _
           " Hex(" & Hex(Err.LastDllError) & ")."
    GetSystemErrorMessageText = "An internal system error occurred with the" & vbCrLf & _
        "FormatMessage API function: " & CStr(Err.LastDllError) & ". No futher information" & vbCrLf & _
        "is available."
    Exit Function
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''
' If FormatMessageResult is not zero, it is the number
' of characters placed in the ErrorText variable.
' Take the left FormatMessageResult characters and
' return that text.
''''''''''''''''''''''''''''''''''''''''''''''''''''''
ErrorText = Left$(ErrorText, FormatMessageResult)
'''''''''''''''''''''''''''''''''''''''''''''
' Get rid of the trailing vbCrLf, if present.
'''''''''''''''''''''''''''''''''''''''''''''
If Len(ErrorText) >= 2 Then
    If Right$(ErrorText, 2) = vbCrLf Then
        ErrorText = Left$(ErrorText, Len(ErrorText) - 2)
    End If
End If

''''''''''''''''''''''''''''''''
' Return the error text as the
' result.
''''''''''''''''''''''''''''''''
GetSystemErrorMessageText = ErrorText

End Function


