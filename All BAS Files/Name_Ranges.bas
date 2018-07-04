Attribute VB_Name = "Name_Ranges"
Option Explicit
Sub Main_Master()
Dim nLastCol As Long, LastRo As Long

'counting variables for loops
Dim i As Integer, j As Integer, k As Integer, t As Integer


'delete named ranges if any
Dim sName As Name
For Each sName In Names
    sName.Delete
Next


nLastCol = Cells(1, Columns.Count).End(xlToLeft).Column 'Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column


Dim myRANGE As Range, MyList As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Create Named Ranges to Build Dynamic Drop Down Lists''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RNGstr As String
Dim N As Name
For i = 1 To nLastCol
'Sheets("Sheet2").Activate

    If Cells(1, i) <> "" Then

        LastRo = Cells(Rows.Count, i).End(xlUp).Row
        Set myRANGE = ActiveSheet.Range(Cells(2, i), Cells(LastRo, i))
        MyList = Cells(1, i).Text
        ActiveSheet.Range(Cells(2, i), Cells(LastRo, i)).Name = MyList
    End If
Next i

End Sub
