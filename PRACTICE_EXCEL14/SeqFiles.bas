Attribute VB_Name = "SeqFiles"
Option Explicit

Sub ReadMe(strFileName As String)
    Dim rLine As String
    Dim i As Integer

    ' line number

    i = 0

    On Error GoTo ExitHere
    Open strFileName For Input As #1

    ' stay inside the loop until the end of file is reached
    Do While Not EOF(1)
       i = i + 1
       Line Input #1, rLine
       MsgBox "Line " & i & " in " & strFileName & " reads: " _
         & Chr(13) & Chr(13) & rLine
    Loop
    MsgBox i & " lines were read."
    Close #1
    Exit Sub
ExitHere:
    MsgBox "File " & strFileName & " could not be found."
End Sub


Sub CountChar(strFileName As String, srchChar As String)
    Dim counter As Integer
    Dim char As String
    
    counter = 0
    Open strFileName For Input As #1

    Do While Not EOF(1)
        char = Input(1, #1)
        If char = srchChar Then
            counter = counter + 1
        End If
    Loop
    If counter <> 0 Then
        MsgBox "Characters (" & srchChar & ") found: " & counter
    Else
        MsgBox "The specified character (" & srchChar & _
          ") has not been found."
    End If
    Close #1
End Sub

Sub ReadAll(strFileName As String)
    Dim all As String
  
    Open strFileName For Input As #1
    all = Input(LOF(1), #1)
    Debug.Print all
    Close #1
End Sub


Sub WriteToTextBox(strFileName As String)
    Dim sh As Worksheet
    Set sh = ActiveWorkbook.Worksheets(3)
  
    On Error GoTo CloseFile:

    Open strFileName For Input As #1
    sh.Shapes.AddTextbox(msoTextOrientationHorizontal, _
      10, 10, 300, 200).Select

    Selection.Characters.Text = Input(LOF(1), #1)
CloseFile:
    Close #1
End Sub

Sub Winners()
    Dim lname As String
    Dim fname As String
    Dim age As Integer

    Open "C:\Excel2013_ByExample\Winners.csv" For Input As #1
    Do While Not EOF(1)
        Input #1, lname, fname, age
        MsgBox lname & ", " & fname & ", " & age
    Loop
    Close #1
End Sub


Sub DataEntry()
    Dim lname As String
    Dim fname As String
    Dim birthdate As Date
    Dim sib As Integer

    Open "C:\Excel2013_ByExample\Friends.txt" For Output As #1
    lname = "Smith"
    fname = "Gregory"
    birthdate = #1/2/1963#
    sib = 3
    Write #1, lname, fname, birthdate, sib

    lname = "Conlin"
    fname = "Janice"
    birthdate = #5/12/1948#
    sib = 1
    Write #1, lname, fname, birthdate, sib

    lname = "Kaufman"
    fname = "Steven"
    birthdate = #4/7/1957#
    sib = 0
    Write #1, lname, fname, birthdate, sib

    Close #1
End Sub

