Attribute VB_Name = "OpenPDFandConvertToWord"
Sub OpenPDFandConvertToWord()
Application.DisplayAlerts = False
Dim wdApp As Object, newDoc As Object
Dim strFile As String
'~~> Establish an Word application object
On Error Resume Next
Set wdApp = CreateObject("Word.Application")
Call wdApp.Run("openpdf.ConvertPDF2Word")
wdApp.Visible = True
Application.DisplayAlerts = True
wdApp = Nothing
    '
    '~~> Rest of the code
    '
End Sub
Sub ImportWordTable()

Dim wdDoc As Object
Dim wdFileName As Variant
Dim tableNo As Integer 'table number in Word
Dim iRow As Long 'row index in Excel
Dim iCol As Integer 'column index in Excel
Dim resultRow As Long
Dim tableStart As Integer
Dim tableTot As Integer
Dim iirow As Long, iicol As Long
On Error Resume Next

ActiveSheet.Range("A:AZ").ClearContents

wdFileName = Application.GetOpenFilename("Word files (*.docx),*.doc", , _
"Browse for file containing table to be imported")

If wdFileName = False Then Exit Sub '(user cancelled import file browser)

Set wdDoc = GetObject(wdFileName) 'open Word file

With wdDoc
wdDoc.Paragraphs.Count
    For Each DocPara In ActiveDocument.Paragraphs

     If Left(DocPara.Range.Style, Len("Heading")) <> "" Then

       Debug.Print DocPara.Range.Text

     End If

    Next

    tableNo = wdDoc.tables.Count
    tableTot = wdDoc.tables.Count
    If tableNo = 0 Then
        MsgBox "This document contains no tables", _
        vbExclamation, "Import Word Table"
    ElseIf tableNo > 1 Then
        tableNo = InputBox("This Word document contains " & tableNo & " tables." & vbCrLf & _
        "Enter the table to start from", "Import Word Table", "1")
    End If

    resultRow = 4

    For tableStart = 1 To tableTot
        With .tables(tableStart)
            'copy cell contents from Word table cells to Excel cells
            For iRow = 1 To .Rows.Count
                For iCol = 1 To .Columns.Count
                    Cells(resultRow, iCol) = WorksheetFunction.Clean(.cell(iRow, iCol).Range.Text)
                Next iCol
                resultRow = resultRow + 1
            Next iRow
        End With
        resultRow = resultRow + 1
    Next tableStart
End With
'cleanImport
End Sub
Sub cleanImport()
With Application
    .ScreenUpdating = False
    .Calculation = xlCalculationManual
    .DisplayAlerts = False
End With
Dim LR As Long: LR = Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Long
For i = LR To 1 Step -1
    If Cells(i, 1).Value = "" Or Cells(i, 1).Value = "/" Then Rows(i).Delete shift:=xlUp
Next i
With Application
    .ScreenUpdating = True
    .Calculation = xlCalculationAutomatic
    .DisplayAlerts = True
End With
End Sub
