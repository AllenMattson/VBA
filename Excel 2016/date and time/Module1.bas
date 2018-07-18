Attribute VB_Name = "Module1"
Option Explicit

Sub DateAndTime()
'   Displays the current date and time
    Dim TheDate As String, TheTime As String
    Dim LastSaved As String
    Dim Greeting As String
    Dim FullName As String, FirstName As String
    Dim SpaceInName As Long
    
    TheDate = Format(Date, "Long Date")
    TheTime = Format(Time, "Medium Time")

'   Determine greeting based on time
    Select Case Time
        Case Is < 0.5:     Greeting = "Good Morning, "
        Case Is >= 0.7083: Greeting = "Good Evening, "
        Case Else:         Greeting = "Good Afternoon, "
    End Select

'   Last saved
    LastSaved = "Document Last Saved: " & ActiveWorkbook.BuiltinDocumentProperties(12)

'   Append user's first name to greeting
    FullName = Application.UserName
    SpaceInName = InStr(1, FullName, " ", 1)
  
'   Handle situation when name has no space
    If SpaceInName = 0 Then SpaceInName = Len(FullName)
    FirstName = Left(FullName, SpaceInName)
    Greeting = Greeting & FirstName

'   Show the message
    MsgBox TheDate & vbCrLf & vbCrLf & "It's " & TheTime, vbOKOnly, Greeting
End Sub

