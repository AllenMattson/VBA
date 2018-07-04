Option Explicit

' Declare for call to mpr.dll.

Declare Function WNetGetUser Lib "mpr.dll" _
        Alias "WNetGetUserA" (ByVal lpName As String, _
                              ByVal lpUserName As String, _
                              lpnlength As Long) As Long

Const NoError = 0       'The Function call was successful

Function GetUserName() 'As String
  If gEnableErrorHandling Then On Error GoTo errHandler

  Const lpnlength As Integer = 255                                      ' Buffer size for the return string
  Dim status As Integer                                                 ' Get return buffer space.
  Dim lpName As String                                                  ' For getting user information.
  Dim lpUserName As String                                              ' For getting user information.

  lpUserName = Space$(lpnlength + 1)                                    ' Assign the buffer size constant to lpUserName.
  
  status = WNetGetUser(lpName, lpUserName, lpnlength)                   ' Get the log-on name of the person using product.

  If status = NoError Then                                              ' See whether error occurred.
    lpUserName = Left$(lpUserName, InStr(lpUserName, Chr(0)) - 1)       ' This line removes the null character. Strings in C are null-
                                                                        '   terminated. Strings in Visual Basic are not null-terminated.
                                                                        '   The null character must be removed from the C strings to be used
                                                                        '   cleanly in Visual Basic.
    GetUserName = lpUserName
  
  Else
    
    GetUserName = "Unknown"
  
  End If


exitHere:
  Exit Function

errHandler:
  MsgBox "Error " & Err.Number & ": " & Err.Description & " in ", _
          vbOKOnly, "Error"

Resume exitHere

End Function
