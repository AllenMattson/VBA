'Public constant must be created to enable global error handling.
'  eg. Public Const gEnableErrorHandling As Boolean = False

Sub Template ()
  
  If gEnableErrorHandling Then On Error GoTo errHandler
  
exitHere:
  
  Exit Function

errHandler:
  MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
          VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"

  Resume exitHere
  
End Function
