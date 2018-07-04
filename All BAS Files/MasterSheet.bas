Attribute VB_Name = "MasterSheet"
Option Explicit
Sub Main_Master()

Dim LR As Integer
Dim wb As Workbook: Set wb = ActiveWorkbook
'Make sure correct workbook is selected
If wb.Name <> "Master.xlsm" Then
Dim msg As String
    msg = MsgBox("Please activate the Master Workbook", vbCritical, "Wrong Workbook")
    Exit Sub
End If
'delete old folder and files already sent to cloud
'Dim TempPathFolder As String
'TempPathFolder = "C:\TempAdvisorFolder"
'KillOldFiles (TempPathFolder)

wb.Activate


Sheets("Sheet1").Activate
Dim nLastCol As Long, LastRo As Long

'counting variables for loops
Dim i As Integer, j As Integer, k As Integer, t As Integer


'delete named ranges if any
Dim sName As Name
For Each sName In Names
    sName.Delete
Next

'delete drop downs in case columns have been changed/updated
With ActiveSheet.Cells.Validation
    .Delete
End With


nLastCol = Sheets("Sheet1").Cells(1, Columns.Count).End(xlToLeft).Column 'Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column



Sheets("Sheet2").Activate

Dim myRANGE As Range, MyList As String
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''Create Named Ranges to Build Dynamic Drop Down Lists''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim RNGstr As String
Dim N As Name
For i = 1 To nLastCol
Sheets("Sheet2").Activate
LastRo = Sheets("Sheet2").Cells(Rows.Count, i).End(xlUp).Row
    If Cells(1, i) <> "" Then
        Set myRANGE = ActiveSheet.Range(Cells(2, i), Cells(LastRo, i))
        MyList = Cells(1, i).Text
        ActiveSheet.Range(Cells(2, i), Cells(LastRo, i)).Name = MyList
    End If
Next i
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub
