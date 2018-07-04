Attribute VB_Name = "OpenPDF"
Option Explicit

'Retrieves a handle to the top-level window whose class name and window name match the specified strings.
'This function does not search child windows. This function does not perform a case-sensitive search.
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
(ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'Retrieves a handle to a window whose class name and window name match the specified strings.
'The function searches child windows, beginning with the one following the specified child window.
'This function does not perform a case-sensitive search.
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
(ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, _
ByVal lpsz2 As String) As Long

'Brings the thread that created the specified window into the foreground and activates the window.
'Keyboard input is directed to the window, and various visual cues are changed for the user.
'The system assigns a slightly higher priority to the thread that created the foreground
'window than it does to other threads.
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

'Sends the specified message to a window or windows. The SendMessage function calls the window procedure
'for the specified window and does not lParenturn until the window procedure has processed the message.
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Places (posts) a message in the message queue associated with the thread that created the specified
'window and lParenturns without waiting for the thread to process the message.
Public Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Constants used in API functions.
Public Const WM_SETTEXT = &HC
Public Const VK_RETURN = &HD
Public Const WM_KEYDOWN = &H100

Private Sub OpenPDF(strPDFPath As String, strPageNumber As String, strZoomValue As String)
    
    'Opens a PDF file to a specific page and with a specific zoom
    'using Adobe Reader Or Adobe Professional.
    'API functions are used to specify the necessary windows
    'and send the page and zoom info to the Adobe window.
    
    'By Christos Samaras
    'http://www.myengineeringworld.net
           
    Dim strPDFName                  As String
    Dim lParent                     As Long
    Dim lFirstChildWindow           As Long
    Dim lSecondChildFirstWindow     As Long
    Dim lSecondChildSecondWindow    As Long
    Dim dtStartTime               As Date
    
    'Check if the PDF path is correct.
    If FileExists(strPDFPath) = False Then
        MsgBox "The PDF path is incorect!", vbCritical, "Wrong path"
        Exit Sub
    End If
    
    'Get the PDF file name from the full path.
    On Error Resume Next
    strPDFName = Mid(strPDFPath, InStrRev(strPDFPath, "\") + 1, Len(strPDFPath))
    On Error GoTo 0
                            
    'The following line depends on the apllication you are using.
    'For Word:
    'ThisDocument.FollowHyperlink strPDFPath, NewWindow:=True
    'For Power Point:
    'ActivePresentation.FollowHyperlink strPDFPath, NewWindow:=True
    'Note that both Word & Power Point pop up a security window asking
    'for access to the specified PDf file.
    'For Access:
    'Application.FollowHyperlink strPDFPath, NewWindow:=True
    'For Excel:
    ThisWorkbook.FollowHyperlink strPDFPath, NewWindow:=True
    'Find the handle of the main/parent window.
    dtStartTime = Now()
    Do Until Now() > dtStartTime + TimeValue("00:00:05")
        lParent = 0
        DoEvents
        'For Adobe Reader.
        'lParent = FindWindow("AcrobatSDIWindow", strPDFName & " - Adobe Reader")
        'For Adobe Professional.
        lParent = FindWindow("AcrobatSDIWindow", strPDFName & " - Adobe Acrobat Pro")
        If lParent <> 0 Then Exit Do
    Loop
    
    If lParent <> 0 Then
    
        'Bring parent window to the foreground (above other windows).
        SetForegroundWindow (lParent)
        
        'Find the handle of the first child window.
        dtStartTime = Now()
        Do Until Now() > dtStartTime + TimeValue("00:00:05")
            lFirstChildWindow = 0
            DoEvents
            lFirstChildWindow = FindWindowEx(lParent, ByVal 0&, vbNullString, "AVUICommandWidget")
            If lFirstChildWindow <> 0 Then Exit Do
        Loop

        'Find the handles of the two subsequent windows.
        If lFirstChildWindow <> 0 Then
            dtStartTime = Now()
            Do Until Now() > dtStartTime + TimeValue("00:00:05")
                lSecondChildFirstWindow = 0
                DoEvents
                lSecondChildFirstWindow = FindWindowEx(lFirstChildWindow, ByVal 0&, "Edit", vbNullString)
                If lSecondChildFirstWindow <> 0 Then Exit Do
            Loop
            
            If lSecondChildFirstWindow <> 0 Then
            
                'Send the zoom value to the corresponding window.
                SendMessage lSecondChildFirstWindow, WM_SETTEXT, 0&, ByVal strZoomValue
                PostMessage lSecondChildFirstWindow, WM_KEYDOWN, VK_RETURN, 0
                
                dtStartTime = Now()
                Do Until Now() > dtStartTime + TimeValue("00:00:05")
                    lSecondChildSecondWindow = 0
                    DoEvents
                    'Notice the difference in syntax between lSecondChildSecondWindow and lSecondChildFirstWindow.
                    'lSecondChildSecondWindow is the handle of the next child window after lSecondChildFirstWindow,
                    'while both windows have as parent window the lFirstChildWindow.
                    lSecondChildSecondWindow = FindWindowEx(lFirstChildWindow, lSecondChildFirstWindow, "Edit", vbNullString)
                    If lSecondChildSecondWindow <> 0 Then Exit Do
                Loop
                If lSecondChildSecondWindow <> 0 Then
                
                    'Send the page number to the corresponding window.
                    SendMessage lSecondChildSecondWindow, WM_SETTEXT, 0&, ByVal strPageNumber
                    PostMessage lSecondChildSecondWindow, WM_KEYDOWN, VK_RETURN, 0
                    
                End If
                
            End If
        
        End If
        
    End If
   
End Sub

Function FileExists(strFilePath As String) As Boolean
    
    'Checks if a file exists.
            
    'By Christos Samaras
    'http://www.myengineeringworld.net

    On Error Resume Next
    If Not Dir(strFilePath, vbDirectory) = vbNullString Then FileExists = True
    On Error GoTo 0
    
End Function

Sub TestPDF()

    OpenPDF ThisWorkbook.Path & "\" & "Sample File.pdf", 6, 143
    
End Sub

