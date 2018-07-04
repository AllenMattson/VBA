Attribute VB_Name = "OpenOutLookOrBindifOpen"
'http://accessexperts.com/blog/2011/07/26/self-healing-objects/
'retrieve an open instance of Outlook or open Outlook if it is closed
#Const LateBind = True

Const olMinimized As Long = 1
Const olMaximized As Long = 2
Const olFolderInbox As Long = 6

#If LateBind Then

Public Function OutlookApp( _
    Optional WindowState As Long = olMinimized, _
    Optional ReleaseIt As Boolean = False _
    ) As Object
    Static o As Object
#Else
Public Function OutlookApp( _
    Optional WindowState As Outlook.OlWindowState = olMinimized, _
    Optional ReleaseIt As Boolean _
) As Outlook.Application
    Static o As Outlook.Application
#End If
On Error GoTo ErrHandler
 
    Select Case True
        Case o Is Nothing, Len(o.Name) = 0
            Set o = GetObject(, "Outlook.Application")
            If o.Explorers.Count = 0 Then
InitOutlook:
                'Open inbox to prevent errors with security prompts
                o.Session.GetDefaultFolder(olFolderInbox).Display
                o.ActiveExplorer.WindowState = WindowState
            End If
        Case ReleaseIt
            Set o = Nothing
    End Select
    Set OutlookApp = o
 
ExitProc:
    Exit Function
ErrHandler:
    Select Case Err.Number
        Case -2147352567
            'User cancelled setup, silently exit
            Set o = Nothing
        Case 429, 462
            Set o = GetOutlookApp()
            If o Is Nothing Then
                Err.Raise 429, "OutlookApp", "Outlook Application does not appear to be installed."
            Else
                Resume InitOutlook
            End If
        Case Else
            MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Unexpected error"
    End Select
    Resume ExitProc
    Resume
End Function

#If LateBind Then
Private Function GetOutlookApp() As Object
#Else
Private Function GetOutlookApp() As Outlook.Application
#End If
On Error GoTo ErrHandler
    
    Set GetOutlookApp = CreateObject("Outlook.Application")
    
ExitProc:
    Exit Function
ErrHandler:
    Select Case Err.Number
        Case Else
            'Do not raise any errors
            Set GetOutlookApp = Nothing
    End Select
    Resume ExitProc
    Resume
End Function

Sub MyMacroThatUseOutlook()
    Dim OutApp  As Object
    Set OutApp = OutlookApp()
    'Automate OutApp as desired
End Sub

