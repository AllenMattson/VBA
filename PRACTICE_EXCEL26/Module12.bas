Attribute VB_Name = "Module12"
Option Explicit

Sub CreateWorkSelChangeEvent()
    Dim objVBCode As CodeModule
    Dim wks As Worksheet
    Dim firstLine As Long

    ' Add a new worksheet
    Set wks = ActiveWorkbook.Worksheets.Add

    ' create a reference to the code module of
    ' the inserted sheet
    Set objVBCode = wks.Parent.VBProject.VBComponents( _
        wks.Name).CodeModule

    ' create an event procedure and return the line at
    ' which the body of the event procedure begins
    
    firstLine = objVBCode.CreateEventProc( _
        "SelectionChange", "Worksheet")

    Debug.Print "Procedure first line: " & firstLine

    ' proceed to add code to the body of the event procedure
    objVBCode.InsertLines firstLine + 1, Chr(9) & _
        "Dim myRange As Range"
    objVBCode.InsertLines firstLine + 2, Chr(9) & _
        "On Error Resume Next"
    objVBCode.InsertLines firstLine + 3, Chr(9) & _
        "Set myRange = Intersect(Range(""A1:A10""),Target)"
    objVBCode.InsertLines firstLine + 4, _
        Chr(9) & "If Not myRange Is Nothing Then"
    objVBCode.InsertLines firstLine + 5, _
        Chr(9) & Chr(9) & _
        "MsgBox ""Data entry or edits are not permitted."""
    objVBCode.InsertLines firstLine + 6, _
        Chr(9) & "End If"

    Set objVBCode = Nothing
    Set wks = Nothing
End Sub



