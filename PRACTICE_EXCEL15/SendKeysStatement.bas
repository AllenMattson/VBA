Attribute VB_Name = "SendKeysStatement"
Option Explicit

Sub FindCPLFiles_Win8()
    ' The keystrokes are for Windows 8
    Shell "Explorer", vbMaximizedFocus

    ' delay the execution by 5 seconds
    Application.Wait (Now + TimeValue("0:00:05"))

    ' Activate the Search box
    SendKeys "{F3}", True
    
    ' delay the execution by 5 seconds
    Application.Wait (Now + TimeValue("0:00:05"))

    ' change the search location to search all folders
    ' on your computer C drive
    SendKeys "%js", True
    SendKeys "%c", True
    SendKeys "%js", True
    SendKeys "%a", True
    ' Activate the Search box
    SendKeys "{F3}", True
    
    ' type in the search string
    SendKeys "*.cpl", True

    ' execute the Search
    SendKeys "{ENTER}", True
    
End Sub

