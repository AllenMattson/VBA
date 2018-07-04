Attribute VB_Name = "NAME_RANGES_AGAIN"
Sub Name_Ranges()
Application.DisplayAlerts = True
Application.Calculation = xlAutomatic
Application.ScreenUpdating = True
Dim nLastCol As Long, LastRo As Long

'counting variables for loops
Dim i As Integer, j As Integer, k As Integer, t As Integer


'delete named ranges if any
Dim sName As Name
For Each sName In Names
    If sName <> "assets" Then sName.Delete
Next

nLastCol = Range("A5").End(xlToRight).Column

Dim myRANGE As Range, MyList As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Create Named Ranges to Build Dynamic Drop Down Lists''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RNGstr As String
Dim n As Name
With ActiveSheet
For i = 5 To nLastCol
'Sheets("Sheet2").Activate

    If Cells(5, i) <> "" Then
        LastRo = Cells(Rows.Count, i).End(xlUp).Row
        Set myRANGE = ActiveSheet.Range(Cells(6, i), Cells(LastRo, i))
        MyList = Cells(5, i).Text
        Range(Cells(6, i), Cells(LastRo, i)).Name = Left(Replace(MyList, " ", ""), 5)
    End If
Next i
End With



End Sub
