Attribute VB_Name = "ImportChild_3"
Sub UpdatedImportChildData()

'delete old sheets if there is any
Dim OldSH As Worksheet
For Each OldSH In ThisWorkbook.Sheets
    If OldSH.Name <> "Sheet1" Then
        If OldSH.Name <> "Sheet2" Then
            If OldSH.Name <> "Sheet3" Then
                Application.DisplayAlerts = False
                OldSH.Delete
                Application.DisplayAlerts = True
            End If
        End If
    End If
Next OldSH
'''''''''''''''''''''''''''''''''''''''''''''''''''
'Child Data Import and use lookup to fill in master
''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim path As String, ThisWB As String, lngFilecounter As Long
Dim wbDest As Workbook, shtDest As Worksheet, ws As Worksheet
Dim FileName As String, Wkb As Workbook
Dim CopyRng As Range, Dest As Range
Dim RowofCopySheet As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''SET PATH TO CHILD FOLDER'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CHECK TO SEE IF PATH NAME ALREADY SET OTHERWISE PROMPT FOR PATH
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets("sheet3").Activate
If Len(Range("B1")) > 2 Then
    path = Sheets("Sheet3").Cells(1, 2).Value
Else
    path = InputBox("Please Insert Folder Path to Child Sheets", "BA Folder Location", ThisWorkbook.path)
    Sheets("Sheet3").Cells(1, 2).Value = path
End If
If path = "" Then
    MsgBox "Insert Path to Import BA sheets"
    Exit Sub
End If
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''CALL FUNCTION TO DETERMINE IF MAC OR PC IS BEING USED''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
If IsMac = False And Right(path, 1) <> "\" Then path = path & "\"
If IsMac = True And Right(path, 1) <> ":" Then path = path & ":"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''******IMPORT THE DATA FROM CHILD SHEETS IF IT EXISTS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   IMPORT INTO 'TEST' SHEET (COPY ALL DATA FROM CHILD, INCLUDING HEADERS FOR NOW)   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets.Add: ActiveSheet.Name = "Test"
RowofCopySheet = 2
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'COMPILES ALL DATA FROM CHILD SHEETS TOGETHER INTO 'TEST' SHEET
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Set shtDest = ActiveWorkbook.Sheets("Test")
FileName = Dir(path & "*.xlsx", vbNormal)
If Len(FileName) = 0 Then Exit Sub
Do Until FileName = vbNullString
    If Not FileName = ThisWB Then
    Application.DisplayAlerts = False
        Set Wkb = Workbooks.Open(FileName:=path & FileName)
        Set CopyRng = Wkb.Sheets(1).Range(Cells(RowofCopySheet, 1), Cells(Cells(Rows.Count, 1).End(xlUp).Row, Cells(1, Columns.Count).End(xlToLeft).Column))
        Set Dest = shtDest.Range("A" & shtDest.Cells(Rows.Count, 1).End(xlUp).Row + 1)
        CopyRng.Copy
        Dest.PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False 'Clear Clipboard
        Wkb.Close False
    End If
FileName = Dir()
Loop
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''REDUNDANT, LEFT HEADERS FOR CHECKING DATA INTEGRITY, LEFT HERE FOR NOW''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''
'Delete any Duplicated Headers
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim k As Integer, LaRo As Integer
LaRo = Cells(Rows.Count, "G").End(xlUp).Row
For k = LaRo To 2 Step -1
    If Trim(Cells(k, 1).Value) = "Date" Then Rows(k).Delete
    If Trim(Cells(k, 1).Value) = "MyDate" Then Rows(k).Delete
Next k
Rows("1:1").Delete

Sheets("Sheet1").Activate
Dim TelRNG As Range, cell As Range
Dim TLR As Integer


If Trim(Range("B2").Value) = "" Then
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NO DATA PREVIOUSLY ENTERED, IMPORT THE BA SHEETS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets("Test").Activate
    Cells(1, 1).CurrentRegion.Copy Sheets("Sheet1").Range("A2")
    Application.DisplayAlerts = False
    Sheets("Test").Delete
    Application.DisplayAlerts = True
    Application.CutCopyMode = False
    Sheets("Sheet1").Activate
    Exit Sub
Else
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Look up Child Data and fill in master
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    LocateMatches
End If
Sheets("Sheet1").Activate
'''''''''''''''''''''''''''''''''''''''''''''
'FORMAT PHONE NUMBERS'
'''''''''''''''''''''''''''''''''''''''''''''
'4/19/17 Below changes implemented per sandeep and brad
TLR = Cells(Rows.Count, "L").End(xlUp).Row
Set TelRNG = Range("L2:L" & TLR)
TelRNG.NumberFormat = "General"
For Each cell In TelRNG
    If Left(LTrim(cell.Value), 1) <> 0 Then cell.Value = " " & 0 & cell.Value
    cell = Replace(Replace(Replace(cell.Value, " ", ""), "-", ""), ".", "")
Next cell
TelRNG.NumberFormat = "0#""-""####""-""####"

'Fix Date Columns
Columns("A:A").NumberFormat = "dd/mm/yy;@"
Columns("AP:AQ").NumberFormat = "dd/mm/yy;@"

''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''
'UNIQUE BUSINESS NAMES Y OR N FOR COLUMN AN ON MASTER
''''''''''''''''''''''''''''''''''''''''''''
Dim Ulr As Integer, Counto As Integer
Ulr = Cells(Rows.Count, "AN").End(xlUp).Row
For Counto = 2 To Ulr
    Cells(Counto, "AN").Select
    ActiveCell.ClearContents
    ActiveCell.FormulaR1C1 = "=IF(COUNTIF(R2C8:R[-6]C[-32],R[-6]C[-32])>1,""N"",""Y"")"
Next
Columns.AutoFit

End Sub
Function IsMac() As Boolean
#If Mac Then
    IsMac = True
#ElseIf Win32 Or Win64 Then
    IsMac = False
#End If
End Function
Sub LocateMatches()
Application.EnableEvents = True

Dim JobNum As String
Dim MyRng As Range
Dim i As Integer, j As Integer
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ASSIGN UNIQUE VALUE IN JOB NUMBER COLUMN FROM SHEET 1 TO VARIABLE JOBNUM
'FIND VARIABLE IN 'TEST' SHEET
'MOVE DATA INPUT BY BA TO MASTER SHEET
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sheets("Sheet1").Activate
Dim LR As Integer: LR = Cells(Rows.Count, 2).End(xlUp).Row
For i = 2 To LR
Sheets("Sheet1").Activate
    JobNum = Range("B" & i).Value 'LOOP TO ASSIGN JOBNUM VARIABLE
    Sheets("Test").Activate
    Dim LRR As Integer: LRR = Cells(Rows.Count, 1).End(xlUp).Row
    For j = 1 To LRR
        If Cells(j, 2).Value = JobNum Then 'FIND VARIABLE
            Range("N" & j & ":AW" & j).Copy Sheets("Sheet1").Range("N" & i) 'MOVE DATA
            Range("AX" & j & ":BB" & j).Copy Sheets("Sheet1").Range("AX" & i) 'MOVE DATA
        Else 'DO SOMETHING ELSE? BACKUP? I DUNNO I'M LEAVING HERE JUST IN CASE
            'next j
        End If
   Next j
   Sheets("Sheet1").Activate
Next i

Application.CutCopyMode = False
'DELETE THE TEST SHEET, THE DATA HAS BEEN SORTED
Application.DisplayAlerts = False
Sheets("Test").Delete
Application.DisplayAlerts = True
End Sub

