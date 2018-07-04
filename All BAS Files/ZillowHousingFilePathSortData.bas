Attribute VB_Name = "ZillowHousingFilePathSortData"
Sub IdentifySources()
Dim WB As Workbook: Set WB = ThisWorkbook
Dim NewWB As Workbook
Dim WS As Worksheet: Set WS = Sheets("Sheet1")
WS.Select
Dim SourceFolder As String
Dim SourceFile As String
Dim SaveFolder As String
Dim SaveFile As String

Dim LR As Integer: LR = Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Integer
With WB
    For i = 3 To LR
        'Get file names
        SourceFolder = Cells(i, 1)
        SourceFolder = SourceFolder & Cells(i, 2)
        SourceFile = Cells(i, 3)
        ''''''''''''''''''''''''''''''''''
        'Identify Where to Save Document
        ''''''''''''''''''''''''''''''''''
        SaveFile = Cells(i, "ZZ").End(xlToLeft).Value
        SaveFolder = Cells(i, "ZZ").End(xlToLeft).Offset(0, -3).Value
        SaveFolder = SaveFolder & Cells(i, "ZZ").End(xlToLeft).Offset(0, -2).Value
        SaveFolder = SaveFolder & Cells(i, "ZZ").End(xlToLeft).Offset(0, -1).Value
        'open source folder and send to macro to create new workbook with changed values
        Set NewWB = Workbooks.Open(SourceFolder & SourceFile)
        '==============================================================================================================
        '==============================================================================================================
        'Run Macro to Change values
        ChangeValues
        '''''
        'Run Macro to Save Document
        '''''
        ActiveWorkbook.SaveAs (SaveFolder & SaveFile)
        ActiveWorkbook.Close True
        '==============================================================================================================
        '==============================================================================================================
        NewWB.Close False
        WB.Activate
    Next i
End With
End Sub
Sub ChangeValues()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False

Dim ZIP As String: ZIP = "RegionName"
Dim MyZip As String

Dim NewWB As Workbook: Set NewWB = ActiveWorkbook
Dim WS As Worksheet: Set WS = ActiveSheet
Dim NewWS As Worksheet
Dim LR As Long: LR = WS.Cells(Rows.Count, 1).End(xlUp).Row
Dim LC As Long: LC = WS.Cells(1, Columns.Count).End(xlToLeft).Column

Dim i As Integer, j As Integer, k As Integer
Dim DateCol As Integer
Dim DateRNG As Range
Set NewWS = Sheets.Add
WS.Activate

'locate first date column then transfer data
    For j = 1 To LC
        If Mid(Cells(1, j), 5, 1) = "-" Then
            DateCol = j
            Cells(1, DateCol).Select
            Set DateRNG = Range(Selection, Selection.End(xlToRight))
            DateRNG.Copy
            NewWS.Range("A2").PasteSpecial xlPasteValues, , , Transpose:=True
            If ActiveSheet.Name <> NewWS.Name Then NewWS.Activate
            GoTo FoundFirstDateColumn
        End If
    Next j


FoundFirstDateColumn:


Cells(1, 1).Value = "Dates"
WS.Activate
'===================================================================
ValuesArray 'Call Macro to Move Data Values to new worksheet (NewWS)
'===================================================================
NewWS.Select
NewWS.Move

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
End Sub
Sub ValuesArray()
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.DisplayAlerts = False

Dim WS As Worksheet: Set WS = ActiveSheet
Dim NewWS As Worksheet
Dim SH As Worksheet
For Each SH In ActiveWorkbook.Worksheets
    If SH.Name <> WS.Name Then Set NewWS = SH
Next SH
WS.Activate








Dim LR As Long: LR = WS.Cells(Rows.Count, 1).End(xlUp).Row
Dim LC As Long: LC = WS.Cells(1, Columns.Count).End(xlToLeft).Column
Dim i As Long, k As Long, DateCol As Long
Dim RNG2 As Range
Dim MyZip As String
Application.CutCopyMode = False


    For k = 2 To LC 'Start at second column so you don't delete the zip codes
        If Mid(Cells(1, k), 5, 1) = "-" Then
            DateCol = k
            For i = LR To 2 Step -1
                If DateCol <= LC Then
                    MyZip = Cells(i, 1).Value
                    If Trim(MyZip) = "" Then GoTo THEWAYOUT
                    MyZip = "ZIP_" & MyZip
                'locate values to transpose
                    Range(Cells(i, DateCol), Cells(i, LC).Offset(0, -1)).Select
                    Set RNG2 = Selection ' Range(Selection, Selection.End(xlToRight)) 'WS.Range(Cells(i, DateCol), Cells(i, LC))
                    
                    RNG2.Copy
                    With NewWS
                        'Paste range values into second row of next empty column
                        NewWS.Activate
                        Range("ZZ1").End(xlToLeft).Offset(0, 1) = MyZip
                        Debug.Print MyZip
                        
                        
                        Range("B900000").End(xlUp).Offset(1, LR - i).Select
                        Selection.PasteSpecial xlPasteValues, , , Transpose:=True

                    End With

                    If ActiveSheet.Name <> NewWS.Name Then NewWS.Activate
                    WS.Activate
                    'delete entire row shifts rows up, count bottom up
                    RNG2.EntireRow.Delete
                    DateCol = DateCol + 1
                End If
            Next i
        Else
            'Delete Column (city,state,metro,county,sizerank)
            Columns(k).EntireColumn.Delete shift:=xlLeft
        End If
    Next k
THEWAYOUT:
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
Application.DisplayAlerts = True
End Sub
Sub FillArrayRange()
Dim WS As Worksheet: Set WS = ActiveSheet
Dim NewWS As Worksheet
'Fill range by transferring array
Dim CellsDown As Long, CellsAcross As Long, StartCol As Integer
Dim i As Long, j As Long, k As Long
Dim TempArray() As Long
Dim TheRange As Range
Dim CurrVal As Long
'get dimensions
CellsDown = Cells(Rows.Count, 1).End(xlUp).Row
If CellsDown = 0 Then Exit Sub

CellsAcross = Cells(1, Columns.Count).End(xlToLeft).Column
For k = 1 To CellsAcross
        If Mid(Cells(1, k), 5, 1) = "-" Then StartCol = k
Next k
'Redimension temp array
ReDim TempArray(2 To CellsDown, StartCol To CellsAcross)
'set worksheet range
Set NewWS = Sheets.Add
Set TheRange = ActiveCell.Range(Cells(1, 1), Cells(CellsDown, CellsAcross))

'Fill temp array
CurrVal = 0
Application.ScreenUpdating = False
For i = 2 To CellsDown
    For j = StartCol To CellsAcross
        TempArray(i, j) = CurrVal + 1
        CurrVal = CurrVal + 1
    Next j
Next i

'Transfer temp array to worksheet
TheRange.Value = TempArray
Application.ScreenUpdating = True
ActiveSheet.Range("A1").End(xlDown).EntireRow.Delete
End Sub
