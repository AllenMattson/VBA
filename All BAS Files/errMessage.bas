Public Sub errMessage(Optional ByVal routineName As String, _
                      Optional ByVal errNumber As String, _
                      Optional ByVal errDescription As String, _
                      Optional ByVal errText As String, _
                      Optional ByVal errLogPath as String = "")
  
  '******************************************************************************
  ' Description:  Writes error message to specifed log file
  '
  ' Author:       taxbender
  ' Contributors:
  ' Sources:
  ' Last Updated: 12/30/2015
  ' Dependencies: Ref - Microsoft Scripting Runtime
  '               Func - checkFolder
  '               Func - padText
  ' Known Issues: None
  '******************************************************************************
  
  On Error GoTo errHandler

  Dim fso As Scripting.File
  Dim errLogFile As String
  Dim errLogMessage As String

  '*** Check for log folder; If not found, create it; If no app folder, exit
  If errLogPath = "" Then
    Application.ActiveWorkbook.Path
  End If
  
  If Not checkFolder(errLogPath) Then
    createFolder errLogPath
  End If

  errLogFile = "errorlog_" & Format$(Now(), "yyyymmdd") & ".log"
  
  '*** Build the log message and add string padding as required

  errLogMessage = Format$(Now(), "mm-dd-yyyy     hh:mm:ss") & Space(5)
  routineName = padText(routineName, 30)
  errNumber = padText(errNumber, 10)
  errDescription = padText(errDescription, 60)
  errLogMessage = errLogMessage & routineName & errNumber & errDescription & errText

  '*** Write error to error log file
  
  Open errLogPath & errLogFile For Append As #1
    Print #1, errLogMessage
  Close #1
  
exitMe:
  Set fso = Nothing
  Exit Sub
  
errHandler:
  Resume exitMe
                        
End Sub
