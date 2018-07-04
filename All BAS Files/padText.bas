Private Function padText(textToPad As String, _
                         textStringLength As Long, _
                         Optional minStringPadding As Long = 5, _
                         Optional padLeftOrRight As String = "Right") As String

  '******************************************************************************
  ' Description:  Adds leading or trailing spaces to a text string
  '
  ' Author:       taxbender
  ' Contributors:
  ' Sources:
  ' Last Updated: 12/30/2015
  ' Dependencies: Var - cEnableErrorHandling
  ' Known Issues: None
  '******************************************************************************

  If cEnableErrorHandling Then On Error GoTo errHandler
  
  Dim textLength As Long
  
  textLength = Len(textToPad)

  If (textLength + minStringPadding) < textStringLength Then
  
    Select Case padLeftOrRight
      
      Case Is = "Left"
        padText = Space(textStringLength - textLength) & textToPad
      
      Case Is = "Right"
        padText = textToPad & Space(textStringLength - textLength)
    End Select
  
  Else
    
    Select Case padLeftOrRight
      Case Is = "Left"
        padText = Space(minStringPadding) & textToPad
      
      Case Is = "Right"
        padText = textToPad & Space(minStringPadding)
    
    End Select
  End If

exitMe:
  Exit Function

errHandler:
  Resume exitMe

End Function
