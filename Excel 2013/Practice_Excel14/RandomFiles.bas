Attribute VB_Name = "RandomFiles"
' create a user-defined data type called Dictionary
Type Dictionary
    en As String * 16 ' English word up to 16 characters
    sp As String * 20 ' Spanish word up to 20 characters
End Type

Sub EnglishToSpanish()
  Dim d As Dictionary
  Dim recNr As Long
  Dim choice As String
  Dim totalRec As Long

  recNr = 1
  ' open the file for random access
  Open "C:\Excel2013_ByExample\Translate.txt" _
      For Random As #1 Len = Len(d)

  Do
    ' get the English word
    choice = InputBox("Enter an English word", "ENGLISH")
    d.en = choice
    ' exit the loop if cancelled
    If choice = "" Then Exit Do
    choice = InputBox("Enter the Spanish equivalent of " _
        & d.en, "SPANISH EQUIVALENT  " & d.en)
    If choice = "" Then Exit Do
    d.sp = choice

    ' write to the record
    Put #1, recNr, d
    ' increase record counter
    recNr = recNr + 1
  'ask for words until Cancel
  Loop Until choice = ""

  totalRec = LOF(1) / Len(d)
  MsgBox "This file contains " & totalRec & " record(s)."
  ' close the file
  Close #1
End Sub


Sub VocabularyDrill()
  Dim d As Dictionary
  Dim totalRec As Long
  Dim recNr As Long
  Dim randomNr As Long
  Dim question As String
  Dim answer As String

  ' open a random access file
  Open "C:\Excel2013_ByExample\Translate.txt" _
      For Random As #1 Len = Len(d)

  ' print the total number of bytes in this file
  Debug.Print "There are " & LOF(1) & " bytes in this file."

  ' find and print the total number of records
  recNr = LOF(1) / Len(d)
  Debug.Print "Total number of records: " & recNr

  Do
    ' get a random record number
    randomNr = Int(recNr * Rnd) + 1
    Debug.Print randomNr

    ' find the random record
    Seek #1, randomNr

    ' read the record
    Get #1, randomNr, d
    Debug.Print Trim(d.en); " "; Trim(d.sp)

    ' assign answer to a variable
    answer = InputBox("What's the Spanish equivalent?", d.en)

    ' finish if cancelled
    If answer = "" Then Close #1: Exit Sub
    Debug.Print answer
        ' check if the answer is correct
        If answer = Trim(d.sp) Then
            MsgBox "Congratulations!"
        Else
            MsgBox "Invalid Answer!!!"
        End If
' keep on asking questions until Cancel is pressed
  Loop While answer <> ""

  ' close file
  Close #1
End Sub


