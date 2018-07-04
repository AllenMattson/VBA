Attribute VB_Name = "TestForWordandOutlook"
Sub TestOutlookIsOpen()
    Dim oOutlook As Object

    On Error Resume Next
    Set oOutlook = GetObject(, "Outlook.Application")
    On Error GoTo 0

    If oOutlook Is Nothing Then
        MsgBox "Outlook is not open, open Outlook and try again"
    Else
        'Call NameOfYourMailMacro
    End If
End Sub
Sub TestMe()
Dim word As Object
Dim doc As Object
On Error Resume Next
Set word = GetObject(, "word.application") 'gives error 429 if Word is not open
If Err.Number = 429 Then
  Err.Clear
  Set word = CreateObject("Word.Application")
End If
If Not word Is Nothing Then
   word.Visible = True
   'Set doc = word.Documents.Add
Else
   MsgBox "Unable to retrieve Word."
End If
End Sub
