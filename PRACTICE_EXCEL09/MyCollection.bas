Attribute VB_Name = "MyCollection"
Option Explicit

Sub GetComments()
  Dim sht As Worksheet
  Dim colNotes As New Collection
  Dim myNote As Comment
  Dim i As Integer
  Dim t As Integer
  Dim strName As String

  strName = InputBox("Enter author's name:")
  For Each sht In ThisWorkbook.Worksheets
    sht.Select
    i = ActiveSheet.Comments.Count
    For Each myNote In ActiveSheet.Comments
        If myNote.Author = strName Then
          MsgBox myNote.Text
          If colNotes.Count = 0 Then
            colNotes.Add Item:=myNote, Key:="first"
          Else
            colNotes.Add Item:=myNote, Before:=1
          End If
      End If
    Next
    t = t + i
  Next
  If colNotes.Count <> 0 Then MsgBox colNotes("first").Text
    MsgBox "Total comments in workbook: " & t & Chr(13) & _
    "Total comments in collection: " & colNotes.Count
    Debug.Print "Comments by " & strName
  For Each myNote In colNotes
    Debug.Print Mid(myNote.Text, Len(myNote.Author) + 2, _
      Len(myNote.Text))
  Next
End Sub

Sub GetComments2()
  Dim sht As Worksheet
  Dim colNotes As New Collection
  Dim myNote As Comment
  Dim i As Integer
  Dim t As Integer
  Dim strName As String
  Dim response
  Dim myID As Integer

  strName = InputBox("Enter author's name:")
  For Each sht In ThisWorkbook.Worksheets
    sht.Select
    i = ActiveSheet.Comments.Count
        For Each myNote In ActiveSheet.Comments
        If myNote.Author = strName Then
          MsgBox myNote.Text
          If colNotes.Count = 0 Then
            colNotes.Add Item:=myNote, Key:="first"
          Else
            colNotes.Add Item:=myNote, Before:=1
          End If
        End If
      Next
    t = t + i
  Next
  If colNotes.Count <> 0 Then MsgBox colNotes("first").Text
     MsgBox "Total comments in workbook: " & t & Chr(13) & _
     "Total comments in collection:" & colNotes.Count
     Debug.Print "Comments by " & strName
     
     myID = 1

     For Each myNote In colNotes
      Debug.Print Mid(myNote.Text, Len(myNote.Author) + 2, _
        Len(myNote.Text))
      response = MsgBox("Remove this comment?" & Chr(13) _
        & Chr(13) & myNote.Text, vbYesNo + vbQuestion)
      If response = 6 Then
        colNotes.Remove index:=myID
      Else
        myID = myID + 1
      End If
     Next

     MsgBox "Total notes in workbook: " & t & Chr(13) & _
     "Total notes in collection: " & colNotes.Count
     Debug.Print "The following comments remain in the collection:"
     
     For Each myNote In colNotes
        Debug.Print Mid(myNote.Text, Len(myNote.Author) + 2, _
        Len(myNote.Text))
     Next
End Sub


Sub DeleteWorkbookComments()
  Dim myComment As Comment
  Dim sht As Worksheet
    
  For Each sht In ThisWorkbook.Worksheets
      For Each myComment In sht.Comments
          myComment.Delete
      Next
  Next
End Sub


